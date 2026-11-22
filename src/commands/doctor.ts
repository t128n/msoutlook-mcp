import { existsSync, statSync } from "node:fs";
import { arch, release } from "node:os";
import { ComError, errorText } from "../com/errors.ts";
import { type AppConfig, DEFAULT_CONFIG, loadConfig, readToken, tokenExists } from "../config/config.ts";
import { Guardrails } from "../config/guardrails.ts";
import {
  DAEMON_TASK,
  OUTLOOK_TASK,
  findOutlookExe,
  newOutlookDefault,
  powershellLanguageMode,
} from "../platform/scheduler.ts";
import type { CommandContext } from "./context.ts";
import { logFileInfo, taskInfoText } from "./lifecycle.ts";

export type CheckLevel = "ok" | "warn" | "fail";

export interface Check {
  level: CheckLevel;
  name: string;
  detail: string;
  fix?: string;
}

export interface DoctorOptions {
  fetchHealth?: (url: string) => Promise<{ status: number; body: unknown }>;
  languageMode?: () => string;
  findOutlookExe?: () => string | null;
  newOutlookDefault?: () => boolean | null;
}

export function printChecks(ctx: CommandContext, checks: Check[]): number {
  for (const check of checks) {
    const label = check.level === "ok" ? "ok  " : check.level === "warn" ? "warn" : "fail";
    ctx.out(`${label}  ${check.name}: ${check.detail}`);
    if (check.fix) ctx.out(`      fix: ${check.fix}`);
  }
  const failed = checks.filter((c) => c.level === "fail").length;
  const warned = checks.filter((c) => c.level === "warn").length;
  ctx.out("");
  ctx.out(`${checks.length} checks, ${failed} failed, ${warned} warnings`);
  return failed > 0 ? 1 : 0;
}

async function defaultFetchHealth(url: string): Promise<{ status: number; body: unknown }> {
  const response = await fetch(url, { signal: AbortSignal.timeout(3_000) });
  return { status: response.status, body: await response.json() };
}

export async function runDoctor(ctx: CommandContext, options: DoctorOptions = {}): Promise<number> {
  const checks: Check[] = [];
  const add = (level: CheckLevel, name: string, detail: string, fix?: string) =>
    checks.push({ level, name, detail, fix });

  // Windows and architecture
  if (ctx.platform === "win32") {
    add("ok", "Windows", `${release()} (${arch()})`);
  } else {
    add("fail", "Windows", `not Windows (${ctx.platform} ${arch()}); this server drives classic Outlook on Windows only`);
  }

  // Config
  let config: AppConfig = { ...DEFAULT_CONFIG };
  let configOk = false;
  try {
    config = loadConfig(ctx.paths);
    configOk = true;
    add("ok", "Config", ctx.paths.configFile);
  } catch (e) {
    const error = e as { message: string; hint?: string };
    add("fail", "Config", error.message, error.hint ?? "run msoutlook-mcp setup");
  }

  // Token
  if (tokenExists(ctx.paths)) {
    add("ok", "Token", ctx.paths.tokenFile);
    if (ctx.platform !== "win32") {
      try {
        const mode = statSync(ctx.paths.tokenFile).mode & 0o777;
        if ((mode & 0o077) !== 0) {
          add("fail", "Token permissions", `mode ${mode.toString(8)} allows other users`, "chmod 600 the token file");
        } else {
          add("ok", "Token permissions", `mode ${mode.toString(8)}`);
        }
      } catch (e) {
        add("warn", "Token permissions", errorText(e));
      }
    } else {
      add("warn", "Token permissions", "Windows ACLs are not inspected by this build; confirm only your account can read the file");
    }
  } else {
    try {
      readToken(ctx.paths);
    } catch (e) {
      const error = e as { message: string; hint?: string };
      add("fail", "Token", error.message, error.hint ?? "run msoutlook-mcp setup");
    }
  }

  // Config file permissions
  if (ctx.platform !== "win32" && configOk) {
    try {
      const mode = statSync(ctx.paths.configFile).mode & 0o777;
      if ((mode & 0o077) !== 0) {
        add("warn", "Config permissions", `mode ${mode.toString(8)} is readable by other users`, "chmod 600 the config file");
      } else {
        add("ok", "Config permissions", `mode ${mode.toString(8)}`);
      }
    } catch {
      /* already reported */
    }
  }

  // Daemon health
  const url = `http://127.0.0.1:${config.port}/health`;
  try {
    const health = await (options.fetchHealth ?? defaultFetchHealth)(url);
    const state = (health.body as { status?: string })?.status ?? "unknown";
    add("ok", "Daemon", `answers on 127.0.0.1:${config.port} (${state})`);
  } catch (e) {
    add("warn", "Daemon", `not answering on 127.0.0.1:${config.port}`, "run msoutlook-mcp install (or serve in a terminal)");
  }

  // Outlook
  const bridge = ctx.bridgeFactory();
  try {
    const ping = await bridge.ping({ timeoutMs: 5_000 });
    add("ok", "Outlook Classic", `${ping.version ?? "version unknown"} (profile ${ping.profile ?? "unknown"})`);
    if (ping.default_store) add("ok", "Default store", ping.default_store);
    else add("warn", "Default store", "not readable", "open Outlook, sign in, and minimize it");
    add("ok", "Mailbox", `inbox ${ping.inbox_count ?? "?"} items, ${ping.inbox_unread ?? "?"} unread`);
  } catch (e) {
    const retryable = e instanceof ComError && e.retryable;
    add(
      retryable ? "warn" : "fail",
      "Outlook Classic",
      errorText(e),
      "start classic Outlook, sign in, then minimize it. Do not close it. The daemon reconnects on its own.",
    );
  } finally {
    bridge.close();
  }

  // New Outlook warning
  if (ctx.platform === "win32") {
    const isNew = (options.newOutlookDefault ?? newOutlookDefault)();
    if (isNew === true) {
      add("warn", "Default Outlook", "the new Outlook appears to be active", "use Outlook Classic; the new Outlook has no COM surface");
    } else if (isNew === false) {
      add("ok", "Default Outlook", "new Outlook preference is off");
    } else {
      add("warn", "Default Outlook", "could not determine from the registry");
    }
  }

  // Exact Outlook executable
  if (ctx.platform === "win32") {
    const exe = (options.findOutlookExe ?? findOutlookExe)();
    add(exe ? "ok" : "warn", "OUTLOOK.EXE", exe ?? "not found in App Paths", exe ? undefined : "repair or reinstall Outlook Classic");
  }

  // PowerShell language mode
  if (ctx.platform === "win32") {
    const mode = (options.languageMode ?? powershellLanguageMode)();
    if (!mode) add("warn", "PowerShell", "could not read the language mode");
    else if (mode === "FullLanguage") add("ok", "PowerShell", mode);
    else add("fail", "PowerShell", mode, "constrained-language mode blocks COM automation; ask IT to allow FullLanguage for PowerShell");
  }

  // Guardrails
  try {
    const status = new Guardrails(config).status();
    const list = status.unrestricted ? "no restriction" : `${status.allowed_recipients.length} entries`;
    if (status.allow_send_direct && status.unrestricted) {
      add("warn", "Guardrails", `sending ON, allowlist ${list}`, "set an allowlist or turn sending off");
    } else {
      add("ok", "Guardrails", `sending ${status.allow_send_direct ? "ON" : "off"}, allowlist ${list}`);
    }
  } catch (e) {
    add("fail", "Guardrails", errorText(e), "fix config.json");
  }

  // Scheduled tasks
  if (ctx.platform === "win32") {
    const daemon = ctx.tasks.query(DAEMON_TASK);
    add(
      daemon.installed ? "ok" : "warn",
      "Daemon task",
      daemon.installed ? taskInfoText(daemon) : "not installed",
      daemon.installed ? undefined : "run msoutlook-mcp install",
    );
    const outlook = ctx.tasks.query(OUTLOOK_TASK);
    add(
      outlook.installed ? "ok" : "warn",
      "Outlook task",
      outlook.installed ? taskInfoText(outlook) : "not installed",
      outlook.installed ? undefined : "run setup and accept the Outlook log-on task",
    );
  }

  // Hidden launcher used by the task
  if (ctx.platform === "win32") {
    add(
      existsSync(ctx.paths.hiddenLauncher) ? "ok" : "warn",
      "Hidden launcher",
      existsSync(ctx.paths.hiddenLauncher) ? ctx.paths.hiddenLauncher : "not found",
      existsSync(ctx.paths.hiddenLauncher) ? undefined : "run msoutlook-mcp install again",
    );
  }

  // Log
  const log = logFileInfo(ctx);
  add("ok", "Log", `${log.path}${log.exists ? ` (${log.bytes} bytes)` : " (not created yet)"}`);

  // Minimize hint, always visible
  add("ok", "Keep Outlook", "minimize Outlook; do not close it. Closing stops the connection.");

  return printChecks(ctx, checks);
}
