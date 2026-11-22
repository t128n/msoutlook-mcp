import type { ComBridge } from "../com/bridge.ts";
import { errorText } from "../com/errors.ts";
import { type AppConfig, DEFAULT_CONFIG, readToken, tokenExists, saveConfig, writeToken } from "../config/config.ts";
import { AllowlistError, compileAllowlist } from "../config/matcher.ts";
import type { AppPaths } from "../config/paths.ts";
import {
  OUTLOOK_TASK,
  type TaskManager,
  currentUser,
  findOutlookExe,
  newOutlookDefault,
  outlookTaskSpec,
} from "../platform/scheduler.ts";

export interface SetupOptions {
  port?: number;
  allowedRecipients?: string[];
  allowSend?: boolean;
  startOutlookAtLogon?: boolean;
  autoLaunchOutlook?: boolean;
  interactive: boolean;
  rotateToken?: boolean;
}

export interface SetupDeps {
  paths: AppPaths;
  platform: NodeJS.Platform;
  bridge: ComBridge;
  tasks: TaskManager;
  out(text: string): void;
  ask(question: string, defaultValue: string): Promise<string>;
  portFree?(port: number): Promise<boolean>;
  findOutlookExe?(): string | null;
  newOutlookDefault?(): boolean | null;
  currentUser?(): string;
}

export const MINIMIZE_HINT =
  "Minimize Outlook; do not close it. Closing Outlook stops the connection. " +
  "To keep it out of the way, right-click the Outlook icon in the notification area and enable " +
  "\"Hide When Minimized\".";

export function mcpSnippet(port: number, token: string): string {
  return JSON.stringify(
    {
      mcpServers: {
        outlook: {
          type: "remote",
          url: `http://127.0.0.1:${port}/mcp`,
          headers: { Authorization: `Bearer ${token}` },
        },
      },
    },
    null,
    2,
  );
}

function yesNo(value: string, fallback: boolean): boolean {
  const v = value.trim().toLowerCase();
  if (!v) return fallback;
  if (["y", "yes", "true", "1", "on"].includes(v)) return true;
  if (["n", "no", "false", "0", "off"].includes(v)) return false;
  return fallback;
}

function splitList(value: string): string[] {
  return value
    .split(/[,;\n]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

async function freePort(start: number, isFree: (port: number) => Promise<boolean>): Promise<number> {
  for (let port = start; port < start + 50 && port <= 65_535; port++) {
    if (await isFree(port)) return port;
  }
  throw new Error(`no free port between ${start} and ${start + 50}`);
}

export async function defaultPortFree(port: number): Promise<boolean> {
  if (port === 0) return true;
  try {
    const server = Bun.serve({ hostname: "127.0.0.1", port, fetch: () => new Response("") });
    server.stop(true);
    return true;
  } catch {
    return false;
  }
}

export async function runSetup(options: SetupOptions, deps: SetupDeps): Promise<number> {
  const { paths, out } = deps;
  const isFree = deps.portFree ?? defaultPortFree;
  const before = DEFAULT_CONFIG;

  out(`msoutlook-mcp setup`);
  out("");

  // 1. Port
  let port = options.port ?? before.port;
  if (options.interactive && options.port === undefined) {
    const answer = await deps.ask(`Port on 127.0.0.1 [${port}]`, String(port));
    const parsed = Number(answer.trim() || port);
    port = Number.isInteger(parsed) && parsed > 0 && parsed <= 65_535 ? parsed : port;
  }
  if (!(await isFree(port))) {
    const chosen = await freePort(port + 1, isFree);
    out(`Port ${port} is in use; using ${chosen} instead.`);
    port = chosen;
  }

  // 2. Allowlist
  let allowed = options.allowedRecipients ?? [];
  if (options.interactive && options.allowedRecipients === undefined) {
    for (;;) {
      const answer = await deps.ask(
        "Allowed recipients (comma separated: address, domain, *@domain, *.domain or /regex/; empty = no restriction)",
        "",
      );
      allowed = splitList(answer);
      try {
        compileAllowlist(allowed);
        break;
      } catch (e) {
        out(`  ${e instanceof AllowlistError ? e.message : String(e)}. Try again.`);
      }
    }
  } else {
    compileAllowlist(allowed);
  }

  // 3. Sending
  let allowSend = options.allowSend ?? false;
  if (options.interactive && options.allowSend === undefined) {
    allowSend = yesNo(await deps.ask("Allow sending mail? Drafts always work; sending is separate. (y/N)", "n"), false);
  }

  // 4. Outlook at log on (Windows only)
  let startOutlookAtLogon = options.startOutlookAtLogon ?? before.start_outlook_at_logon;
  if (deps.platform !== "win32") {
    startOutlookAtLogon = false;
  } else if (options.interactive && options.startOutlookAtLogon === undefined) {
    startOutlookAtLogon = yesNo(
      await deps.ask("Start classic Outlook minimized at log on? (Y/n)", "y"),
      true,
    );
  }

  // 5. Auto-launch
  let autoLaunchOutlook = options.autoLaunchOutlook ?? before.auto_launch_outlook;
  if (options.interactive && options.autoLaunchOutlook === undefined) {
    autoLaunchOutlook = yesNo(
      await deps.ask(
        "Let the daemon start Outlook when it is closed? Off is recommended; the logon task already starts it. (y/N)",
        "n",
      ),
      false,
    );
  }

  const config: AppConfig = {
    ...before,
    port,
    allow_send: allowSend,
    allowed_recipients: allowed,
    start_outlook_at_logon: startOutlookAtLogon,
    auto_launch_outlook: autoLaunchOutlook,
  };

  try {
    saveConfig(paths, config);
  } catch (e) {
    out(`Could not write ${paths.configFile}: ${errorText(e)}`);
    return 1;
  }

  let token: string;
  try {
    token = tokenExists(paths) && !options.rotateToken ? "" : writeToken(paths);
  } catch (e) {
    out(`Could not write the access token: ${errorText(e)}`);
    return 1;
  }
  if (!token) token = readToken(paths);

  out("");
  out(`Config written to ${paths.configFile}`);
  out(`Access token written to ${paths.tokenFile} (readable by you only)`);
  out("");

  // 6. Check Outlook, but never make setup fail because it is not running.
  out("Outlook:");
  try {
    const ping = await deps.bridge.ping({ timeoutMs: 5_000 });
    out(`  OK: classic Outlook ${ping.version ?? "(version unknown)"}, profile ${ping.profile ?? "(unknown)"}`);
  } catch (e) {
    out(`  Not reachable yet: ${errorText(e)}`);
    out("  Start classic Outlook, sign in, then minimize it.");
  }
  const isNew = (deps.newOutlookDefault ?? newOutlookDefault)();
  if (isNew === true) {
    out("  Warning: the new Outlook appears to be the default. This server needs Outlook Classic (no COM surface in the new one).");
  }

  // 7. Outlook logon task
  if (deps.platform === "win32" && startOutlookAtLogon) {
    try {
      const exe = (deps.findOutlookExe ?? findOutlookExe)();
      if (!exe) {
        out(`  Could not find OUTLOOK.EXE in the App Paths registry key; no Outlook logon task registered.`);
      } else {
        const user = (deps.currentUser ?? currentUser)();
        deps.tasks.create(OUTLOOK_TASK, outlookTaskSpec(exe, user));
        out(`  Registered the logon task "${OUTLOOK_TASK}" for ${user} (starts Outlook minimized).`);
      }
    } catch (e) {
      out(`  Warning: could not register the Outlook logon task: ${errorText(e)}`);
    }
  }

  out("");
  out(MINIMIZE_HINT);
  out("To set it: right-click the Outlook icon in the notification area and click \"Hide When Minimized\".");
  out("");
  out("Add this to your MCP client:");
  out("");
  out(mcpSnippet(port, token));
  out("");
  out("Next:");
  out("  msoutlook-mcp install    # copy the exe and register the daemon logon task");
  out("  msoutlook-mcp doctor     # check everything");
  return 0;
}
