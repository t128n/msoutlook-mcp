import { randomBytes } from "node:crypto";
import { readFileSync, rmSync, statSync } from "node:fs";
import { compileAllowlist, AllowlistError } from "./matcher.ts";
import { type AppPaths, ensurePrivateDir, writePrivateFile } from "./paths.ts";

export const CONFIG_VERSION = 1;

export type LogLevel = "debug" | "info" | "warn" | "error";

export interface AppConfig {
  version: 1;
  port: number;
  allow_send: boolean;
  allowed_recipients: string[];
  shared_mailboxes: string[];
  auto_launch_outlook: boolean;
  start_outlook_at_logon: boolean;
  poll_interval_ms: number;
  event_folders: string[];
  log_level: LogLevel;
}

export const DEFAULT_CONFIG: AppConfig = {
  version: CONFIG_VERSION,
  port: 28471,
  allow_send: false,
  allowed_recipients: [],
  shared_mailboxes: [],
  auto_launch_outlook: false,
  start_outlook_at_logon: true,
  poll_interval_ms: 30_000,
  event_folders: ["inbox"],
  log_level: "info",
};

export class ConfigError extends Error {
  readonly hint: string | undefined;
  constructor(message: string, hint?: string) {
    super(message);
    this.name = "ConfigError";
    this.hint = hint;
  }
}

const object = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new ConfigError("config must be a JSON object");
  }
  return value as Record<string, unknown>;
};

const boolean = (value: unknown, key: string, fallback: boolean): boolean => {
  if (value === undefined) return fallback;
  if (typeof value !== "boolean") throw new ConfigError(`config.${key} must be true or false`);
  return value;
};

const stringArray = (value: unknown, key: string, fallback: string[]): string[] => {
  if (value === undefined) return [...fallback];
  if (!Array.isArray(value) || value.some((v) => typeof v !== "string")) {
    throw new ConfigError(`config.${key} must be an array of strings`);
  }
  return (value as string[]).map((v) => v.trim()).filter(Boolean);
};

export function validateConfig(input: unknown): AppConfig {
  const raw = object(input);
  const version = raw.version === undefined ? CONFIG_VERSION : raw.version;
  if (version !== CONFIG_VERSION) {
    throw new ConfigError(`unsupported config version ${String(version)}`, "Update msoutlook-mcp or run setup again.");
  }

  const port = raw.port === undefined ? DEFAULT_CONFIG.port : raw.port;
  if (typeof port !== "number" || !Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new ConfigError("config.port must be an integer between 1 and 65535");
  }

  const poll = raw.poll_interval_ms === undefined ? DEFAULT_CONFIG.poll_interval_ms : raw.poll_interval_ms;
  if (typeof poll !== "number" || !Number.isInteger(poll) || poll < 1_000 || poll > 3_600_000) {
    throw new ConfigError("config.poll_interval_ms must be between 1000 and 3600000");
  }

  const level = raw.log_level ?? DEFAULT_CONFIG.log_level;
  if (level !== "debug" && level !== "info" && level !== "warn" && level !== "error") {
    throw new ConfigError("config.log_level must be debug, info, warn or error");
  }

  const allowed = stringArray(raw.allowed_recipients, "allowed_recipients", DEFAULT_CONFIG.allowed_recipients);
  try {
    compileAllowlist(allowed);
  } catch (e) {
    if (e instanceof AllowlistError) throw new ConfigError(e.message, "Fix the entry or remove it, then retry.");
    throw e;
  }

  return {
    version: CONFIG_VERSION,
    port,
    allow_send: boolean(raw.allow_send, "allow_send", DEFAULT_CONFIG.allow_send),
    allowed_recipients: allowed,
    shared_mailboxes: stringArray(raw.shared_mailboxes, "shared_mailboxes", DEFAULT_CONFIG.shared_mailboxes).map(
      (v) => v.toLowerCase(),
    ),
    auto_launch_outlook: boolean(raw.auto_launch_outlook, "auto_launch_outlook", DEFAULT_CONFIG.auto_launch_outlook),
    start_outlook_at_logon: boolean(
      raw.start_outlook_at_logon,
      "start_outlook_at_logon",
      DEFAULT_CONFIG.start_outlook_at_logon,
    ),
    poll_interval_ms: poll,
    event_folders: stringArray(raw.event_folders, "event_folders", DEFAULT_CONFIG.event_folders),
    log_level: level,
  };
}

export function loadConfig(paths: AppPaths): AppConfig {
  let text: string;
  try {
    text = readFileSync(paths.configFile, "utf8");
  } catch {
    throw new ConfigError(`no config at ${paths.configFile}`, "Run msoutlook-mcp setup first.");
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new ConfigError(`config is not valid JSON: ${e instanceof Error ? e.message : String(e)}`);
  }
  return validateConfig(parsed);
}

export function saveConfig(paths: AppPaths, config: AppConfig): void {
  ensurePrivateDir(paths.configDir);
  writePrivateFile(paths.configFile, JSON.stringify(config, null, 2) + "\n");
}

export function hasConfig(paths: AppPaths): boolean {
  try {
    statSync(paths.configFile);
    return true;
  } catch {
    return false;
  }
}

export function generateToken(): string {
  return randomBytes(32).toString("base64url");
}

/** Writes a new token, mode 0600, and returns it. */
export function writeToken(paths: AppPaths): string {
  const token = generateToken();
  ensurePrivateDir(paths.configDir);
  writePrivateFile(paths.tokenFile, token + "\n");
  return token;
}

export function readToken(paths: AppPaths): string {
  let text: string;
  try {
    text = readFileSync(paths.tokenFile, "utf8").trim();
  } catch {
    throw new ConfigError(`no access token at ${paths.tokenFile}`, "Run msoutlook-mcp setup first.");
  }
  if (!text) throw new ConfigError(`access token at ${paths.tokenFile} is empty`, "Run msoutlook-mcp setup --force.");
  return text;
}

export function tokenExists(paths: AppPaths): boolean {
  try {
    return readFileSync(paths.tokenFile, "utf8").trim().length > 0;
  } catch {
    return false;
  }
}

export function purgeConfigAndState(paths: AppPaths): void {
  rmSync(paths.configFile, { force: true });
  rmSync(paths.tokenFile, { force: true });
  rmSync(paths.stateFile, { force: true });
  rmSync(paths.eventsFile, { force: true });
}
