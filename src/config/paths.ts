import { chmodSync, existsSync, mkdirSync, renameSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

export interface AppPaths {
  home: string;
  configDir: string;
  stateDir: string;
  configFile: string;
  tokenFile: string;
  stateFile: string;
  eventsFile: string;
  logDir: string;
  logFile: string;
  installDir: string;
  installedExe: string;
  hiddenLauncher: string;
}

const NAME = "msoutlook-mcp";

const first = (...values: (string | undefined)[]): string | undefined =>
  values.find((v) => v !== undefined && v !== "");

/**
 * Windows keeps everything under %LOCALAPPDATA%\msoutlook-mcp.
 * Elsewhere config and state follow XDG, and MSOUTLOOK_MCP_HOME overrides both
 * with one directory (tests use this).
 */
export function appPaths(
  env: Record<string, string | undefined> = process.env,
  platform: NodeJS.Platform = process.platform,
): AppPaths {
  const override = first(env.MSOUTLOOK_MCP_HOME);
  let configDir: string;
  let stateDir: string;
  let home: string;

  if (platform === "win32") {
    home = override ?? join(first(env.LOCALAPPDATA, env.APPDATA, homedir())!, NAME);
    configDir = home;
    stateDir = home;
  } else if (override) {
    home = override;
    configDir = override;
    stateDir = override;
  } else {
    const configBase = first(env.XDG_CONFIG_HOME, join(homedir(), ".config"))!;
    const stateBase = first(env.XDG_STATE_HOME, join(homedir(), ".local", "state"))!;
    configDir = join(configBase, NAME);
    stateDir = join(stateBase, NAME);
    home = configDir;
  }

  const installDir = platform === "win32" ? home : join(stateDir, "install");
  return {
    home,
    configDir,
    stateDir,
    configFile: join(configDir, "config.json"),
    tokenFile: join(configDir, "token"),
    stateFile: join(stateDir, "state.json"),
    eventsFile: join(stateDir, "events.json"),
    logDir: join(stateDir, "logs"),
    logFile: join(stateDir, "logs", "server.log"),
    installDir,
    installedExe: join(installDir, platform === "win32" ? "msoutlook-mcp.exe" : "msoutlook-mcp"),
    hiddenLauncher: join(installDir, "start-hidden.vbs"),
  };
}

/** Best effort on Windows: `%LOCALAPPDATA%` is already per user, but keep the ACL tight. */
function hardenWindowsAcl(path: string, directory: boolean): void {
  if (process.platform !== "win32") return;
  const domain = process.env.USERDOMAIN;
  const name = process.env.USERNAME;
  const user = domain && name ? `${domain}\\${name}` : name;
  if (!user) return;
  try {
    Bun.spawnSync(
      [
        "icacls.exe",
        path,
        "/inheritance:r",
        "/grant:r",
        directory ? `${user}:(OI)(CI)F` : `${user}:F`,
      ],
      { stdout: "ignore", stderr: "ignore", windowsHide: true },
    );
  } catch {
    /* not fatal: the file is still inside the user's profile */
  }
}

export function ensurePrivateDir(dir: string): void {
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  if (process.platform === "win32") {
    hardenWindowsAcl(dir, true);
    return;
  }
  try {
    chmodSync(dir, 0o700);
  } catch {
    /* a volume that does not support modes */
  }
}

/** Atomic write, mode 0600. The token and config are for this user only. */
export function writePrivateFile(file: string, contents: string): void {
  ensurePrivateDir(dirname(file));
  const temporary = `${file}.tmp-${process.pid}`;
  writeFileSync(temporary, contents, { mode: 0o600 });
  if (process.platform !== "win32") chmodSync(temporary, 0o600);
  renameSync(temporary, file);
  if (process.platform === "win32") {
    hardenWindowsAcl(file, false);
  } else {
    try {
      chmodSync(file, 0o600);
    } catch {
      /* ignore */
    }
  }
}

export function fileExists(file: string): boolean {
  return existsSync(file);
}
