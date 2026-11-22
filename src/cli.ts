import { createInterface } from "node:readline/promises";
import { PowerShellComBridge } from "./com/index.ts";
import { ConfigError } from "./config/config.ts";
import { appPaths } from "./config/paths.ts";
import { runConfig } from "./commands/config-command.ts";
import type { CommandContext } from "./commands/context.ts";
import { runDoctor } from "./commands/doctor.ts";
import { runInstall, runUninstall, runUpgrade } from "./commands/lifecycle.ts";
import { runServe } from "./commands/serve.ts";
import { runSetup } from "./commands/setup.ts";
import { isCompiled } from "./platform/index.ts";
import { run, windowsTaskManager } from "./platform/scheduler.ts";
import { VERSION } from "./version.ts";

export interface FlagSpec {
  type: "boolean" | "string";
  description: string;
  arg?: string;
}

export type FlagSpecs = Record<string, FlagSpec>;

export interface ParsedFlags {
  values: Record<string, string | boolean>;
  positionals: string[];
  help: boolean;
}

export class UsageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UsageError";
  }
}

export function parseFlags(argv: string[], specs: FlagSpecs): ParsedFlags {
  const values: Record<string, string | boolean> = {};
  const positionals: string[] = [];
  let help = false;

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!;
    if (arg === "-h" || arg === "--help") {
      help = true;
      continue;
    }
    if (arg === "--") {
      positionals.push(...argv.slice(i + 1));
      break;
    }
    if (arg.startsWith("--")) {
      const body = arg.slice(2);
      const equals = body.indexOf("=");
      let name = equals >= 0 ? body.slice(0, equals) : body;
      let inline = equals >= 0 ? body.slice(equals + 1) : undefined;
      let negated = false;
      if (!specs[name] && name.startsWith("no-")) {
        const candidate = name.slice(3);
        if (specs[candidate]?.type === "boolean") {
          name = candidate;
          negated = true;
        }
      }
      const spec = specs[name];
      if (!spec) throw new UsageError(`unknown option --${name}`);
      if (spec.type === "boolean") {
        if (negated) values[name] = false;
        else if (inline !== undefined) values[name] = inline !== "false" && inline !== "0" && inline !== "off";
        else values[name] = true;
      } else {
        if (inline === undefined) {
          i += 1;
          if (i >= argv.length) throw new UsageError(`--${name} needs ${spec.arg ?? "a value"}`);
          inline = argv[i]!;
        }
        values[name] = inline;
      }
      continue;
    }
    if (arg.startsWith("-") && arg !== "-") throw new UsageError(`unknown option ${arg}`);
    positionals.push(arg);
  }
  return { values, positionals, help };
}

export function usage(command: string, specs: FlagSpecs): string {
  const lines = [`Usage: msoutlook-mcp ${command} [options]`];
  if (Object.keys(specs).length > 0) {
    lines.push("", "Options:");
    for (const [name, spec] of Object.entries(specs)) {
      const arg = spec.type === "string" ? ` <${spec.arg ?? "value"}>` : "";
      lines.push(`  --${name}${arg}`.padEnd(34) + spec.description);
    }
    lines.push(`  --help`.padEnd(34) + "show this help");
  }
  return lines.join("\n");
}

export const GENERAL_HELP = `msoutlook-mcp ${VERSION}

An MCP server for classic Outlook on Windows through COM, with polling-based
new_mail events. Minimize Outlook; do not close it.

Usage: msoutlook-mcp <command> [options]

Commands:
  setup       first-time configuration: config, token, allowlist, MCP snippet
  doctor      check every prerequisite; one line per check with the fix
  install     copy the exe and register the daemon logon task
  uninstall   stop and remove the tasks; keep config unless --purge
  upgrade     run the new exe as: <new>.exe upgrade
  serve       run the daemon in the foreground (the logon task runs this)
  config      show or change the config without restarting
  version     print the version
  --help      show this help

Run "msoutlook-mcp <command> --help" for command options.`;

export const SETUP_FLAGS: FlagSpecs = {
  port: { type: "string", arg: "port", description: "port on 127.0.0.1 (default 28471)" },
  allow: { type: "string", arg: "list", description: "allowed recipients, comma separated" },
  "allow-send": { type: "boolean", description: "allow sending (default off)" },
  "outlook-logon": { type: "boolean", description: "start Outlook minimized at log on" },
  "auto-launch-outlook": { type: "boolean", description: "let the daemon start Outlook (default off)" },
  "rotate-token": { type: "boolean", description: "generate a new access token" },
  "non-interactive": { type: "boolean", description: "use flags and defaults, ask nothing" },
};

export const INSTALL_FLAGS: FlagSpecs = {
  "no-start": { type: "boolean", description: "register the task but do not start it now" },
};

export const UNINSTALL_FLAGS: FlagSpecs = {
  purge: { type: "boolean", description: "also remove config, token, state and logs" },
};

export const UPGRADE_FLAGS: FlagSpecs = {
  force: { type: "boolean", description: "allow a downgrade" },
};

export const SERVE_FLAGS: FlagSpecs = {
  port: { type: "string", arg: "port", description: "listen on this port instead of config.port" },
  console: { type: "boolean", description: "also write log lines to the console" },
};

function flag(options: Record<string, string | boolean>, name: string): boolean | undefined {
  const value = options[name];
  return typeof value === "boolean" ? value : undefined;
}

function stringFlag(options: Record<string, string | boolean>, name: string): string | undefined {
  const value = options[name];
  return typeof value === "string" ? value : undefined;
}

function defaultContext(overrides: Partial<CommandContext> = {}): CommandContext {
  const platform = overrides.platform ?? process.platform;
  const env = overrides.env ?? process.env;
  return {
    paths: overrides.paths ?? appPaths(env, platform),
    platform,
    env,
    execPath: overrides.execPath ?? process.execPath,
    compiled: overrides.compiled ?? isCompiled(overrides.execPath ?? process.execPath),
    isTTY: overrides.isTTY ?? Boolean(process.stdin.isTTY),
    out: overrides.out ?? ((text?: string) => console.log(text ?? "")),
    err: overrides.err ?? ((text?: string) => console.error(text ?? "")),
    bridgeFactory: overrides.bridgeFactory ?? (() => new PowerShellComBridge()),
    tasks: overrides.tasks ?? windowsTaskManager,
    now: overrides.now ?? (() => new Date()),
    run: overrides.run ?? run,
  };
}

export async function main(argv: string[], overrides: Partial<CommandContext> = {}): Promise<number> {
  const ctx = defaultContext(overrides);
  const command = argv[0];

  if (!command || command === "--help" || command === "-h" || command === "help") {
    ctx.out(GENERAL_HELP);
    return 0;
  }
  if (command === "version" || command === "--version") {
    ctx.out(VERSION);
    return 0;
  }

  try {
    switch (command) {
      case "setup": {
        const parsed = parseFlags(argv.slice(1), SETUP_FLAGS);
        if (parsed.help) {
          ctx.out(usage("setup", SETUP_FLAGS));
          return 0;
        }
        if (parsed.positionals.length > 0) throw new UsageError(`unexpected argument "${parsed.positionals[0]}"`);
        const nonInteractive = flag(parsed.values, "non-interactive") === true;
        const interactive = !nonInteractive && ctx.isTTY;
        const allow = stringFlag(parsed.values, "allow");
        const portText = stringFlag(parsed.values, "port");
        const port = portText === undefined ? undefined : Number(portText);
        if (port !== undefined && (!Number.isInteger(port) || port < 1 || port > 65_535)) {
          throw new UsageError("--port must be an integer between 1 and 65535");
        }
        let rl: ReturnType<typeof createInterface> | undefined;
        try {
          if (interactive) {
            rl = createInterface({ input: process.stdin, output: process.stdout });
            const ask = async (question: string, defaultValue: string): Promise<string> => {
              const answer = await rl!.question(`${question}: `);
              return answer.trim() || defaultValue;
            };
            return await runSetup(
              {
                port,
                allowedRecipients: allow === undefined ? undefined : allow.split(/[,;]/).map((v) => v.trim()).filter(Boolean),
                allowSend: flag(parsed.values, "allow-send"),
                startOutlookAtLogon: flag(parsed.values, "outlook-logon"),
                autoLaunchOutlook: flag(parsed.values, "auto-launch-outlook"),
                rotateToken: flag(parsed.values, "rotate-token") === true,
                interactive: true,
              },
              {
                paths: ctx.paths,
                platform: ctx.platform,
                bridge: ctx.bridgeFactory(),
                tasks: ctx.tasks,
                out: ctx.out,
                ask,
              },
            );
          }
          return await runSetup(
            {
              port,
              allowedRecipients: allow === undefined ? undefined : allow.split(/[,;]/).map((v) => v.trim()).filter(Boolean),
              allowSend: flag(parsed.values, "allow-send"),
              startOutlookAtLogon: flag(parsed.values, "outlook-logon"),
              autoLaunchOutlook: flag(parsed.values, "auto-launch-outlook"),
              rotateToken: flag(parsed.values, "rotate-token") === true,
              interactive: false,
            },
            {
              paths: ctx.paths,
              platform: ctx.platform,
              bridge: ctx.bridgeFactory(),
              tasks: ctx.tasks,
              out: ctx.out,
              ask: async (_question, defaultValue) => defaultValue,
            },
          );
        } finally {
          rl?.close();
        }
      }

      case "doctor": {
        const parsed = parseFlags(argv.slice(1), {});
        if (parsed.help) {
          ctx.out(usage("doctor", {}));
          return 0;
        }
        return await runDoctor(ctx);
      }

      case "install": {
        const parsed = parseFlags(argv.slice(1), INSTALL_FLAGS);
        if (parsed.help) {
          ctx.out(usage("install", INSTALL_FLAGS));
          return 0;
        }
        return runInstall(ctx, { noStart: flag(parsed.values, "no-start") === true });
      }

      case "uninstall": {
        const parsed = parseFlags(argv.slice(1), UNINSTALL_FLAGS);
        if (parsed.help) {
          ctx.out(usage("uninstall", UNINSTALL_FLAGS));
          return 0;
        }
        return runUninstall(ctx, { purge: flag(parsed.values, "purge") === true });
      }

      case "upgrade": {
        const parsed = parseFlags(argv.slice(1), UPGRADE_FLAGS);
        if (parsed.help) {
          ctx.out(usage("upgrade", UPGRADE_FLAGS));
          return 0;
        }
        return runUpgrade(ctx, { force: flag(parsed.values, "force") === true });
      }

      case "serve": {
        const parsed = parseFlags(argv.slice(1), SERVE_FLAGS);
        if (parsed.help) {
          ctx.out(usage("serve", SERVE_FLAGS));
          return 0;
        }
        const portText = stringFlag(parsed.values, "port");
        const port = portText === undefined ? undefined : Number(portText);
        if (port !== undefined && (!Number.isInteger(port) || port < 1 || port > 65_535)) {
          throw new UsageError("--port must be an integer between 1 and 65535");
        }
        return await runServe(ctx, { port, console: flag(parsed.values, "console") === true });
      }

      case "config": {
        const parsed = parseFlags(argv.slice(1), {});
        if (parsed.help) {
          ctx.out(usage("config", {}));
          return 0;
        }
        return runConfig(ctx, [parsed.positionals[0] ?? "show", ...parsed.positionals.slice(1)]);
      }

      default:
        ctx.err(`error: unknown command "${command}"`);
        ctx.err(GENERAL_HELP);
        return 2;
    }
  } catch (e) {
    if (e instanceof UsageError) {
      ctx.err(`error: ${e.message}`);
      return 2;
    }
    if (e instanceof ConfigError) {
      ctx.err(`error: ${e.message}`);
      if (e.hint) ctx.err(`fix: ${e.hint}`);
      return 1;
    }
    ctx.err(`error: ${e instanceof Error ? e.message : String(e)}`);
    return 1;
  }
}
