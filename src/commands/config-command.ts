import { type AppConfig, loadConfig, saveConfig, validateConfig } from "../config/config.ts";
import { Guardrails } from "../config/guardrails.ts";
import { compileAllowlist } from "../config/matcher.ts";
import type { CommandContext } from "./context.ts";

const BOOLEAN_KEYS: Record<string, keyof AppConfig> = {
  "allow-send": "allow_send",
  "auto-launch-outlook": "auto_launch_outlook",
  "outlook-logon": "start_outlook_at_logon",
};

const LIST_KEYS: Record<string, keyof AppConfig> = {
  "allowed-recipients": "allowed_recipients",
  "shared-mailboxes": "shared_mailboxes",
  "event-folders": "event_folders",
};

export const CONFIG_USAGE = `Usage:
  msoutlook-mcp config show
  msoutlook-mcp config set <key> <value>

Keys:
  allow-send on|off                 read at send time, no restart
  allowed-recipients <list>         comma separated; empty string clears
  shared-mailboxes <list>           addresses allowed for from_mailbox
  event-folders <list>              folders polled by new_mail
  port <1-65535>
  poll-interval <seconds>
  auto-launch-outlook on|off
  outlook-logon on|off
  log-level debug|info|warn|error`;

function parseBoolean(value: string): boolean {
  const v = value.trim().toLowerCase();
  if (["on", "true", "yes", "1"].includes(v)) return true;
  if (["off", "false", "no", "0"].includes(v)) return false;
  throw new Error(`expected on or off, got "${value}"`);
}

function splitList(value: string): string[] {
  return value
    .split(/[,;\n]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

export function runConfig(ctx: CommandContext, args: string[]): number {
  const action = args[0];
  if (!action || action === "--help" || action === "-h" || action === "help") {
    ctx.out(CONFIG_USAGE);
    return action ? 0 : 1;
  }

  let config: AppConfig;
  try {
    config = loadConfig(ctx.paths);
  } catch (e) {
    ctx.err(`error: ${(e as Error).message}`);
    const hint = (e as { hint?: string }).hint;
    if (hint) ctx.err(`fix: ${hint}`);
    return 1;
  }

  if (action === "show") {
    const status = new Guardrails(config).status();
    ctx.out(
      JSON.stringify(
        {
          config: { ...config, token: "(stored separately, never printed)" },
          guardrails: status,
          files: { config: ctx.paths.configFile, token: ctx.paths.tokenFile },
        },
        null,
        2,
      ),
    );
    return 0;
  }

  if (action !== "set") {
    ctx.err(`error: unknown config action "${action}"`);
    ctx.err(CONFIG_USAGE);
    return 1;
  }

  const key = args[1];
  const raw = args[2];
  if (!key || raw === undefined) {
    ctx.err("error: config set needs a key and a value");
    ctx.err(CONFIG_USAGE);
    return 1;
  }

  try {
    if (BOOLEAN_KEYS[key]) {
      (config as unknown as Record<string, unknown>)[BOOLEAN_KEYS[key]] = parseBoolean(raw);
    } else if (LIST_KEYS[key]) {
      const list = splitList(raw);
      if (key === "allowed-recipients") compileAllowlist(list);
      (config as unknown as Record<string, unknown>)[LIST_KEYS[key]] = list;
    } else if (key === "port") {
      const port = Number(raw);
      if (!Number.isInteger(port)) throw new Error("port must be an integer");
      config.port = port;
    } else if (key === "poll-interval") {
      const seconds = Number(raw);
      if (!Number.isFinite(seconds)) throw new Error("poll-interval must be a number of seconds");
      config.poll_interval_ms = Math.round(seconds * 1000);
    } else if (key === "log-level") {
      if (!["debug", "info", "warn", "error"].includes(raw)) {
        throw new Error("log-level must be debug, info, warn or error");
      }
      config.log_level = raw as AppConfig["log_level"];
    } else {
      throw new Error(`unknown key "${key}"`);
    }

    const validated = validateConfig(config);
    saveConfig(ctx.paths, validated);
    ctx.out(`Set ${key}.`);
    if (key === "allow-send") {
      ctx.out(`Sending is now ${validated.allow_send ? "ON" : "off"}. This is read at send time; restarting is not needed.`);
    }
    return 0;
  } catch (e) {
    ctx.err(`error: ${(e as Error).message}`);
    return 1;
  }
}
