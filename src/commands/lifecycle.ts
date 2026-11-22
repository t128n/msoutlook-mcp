import { copyFileSync, existsSync, mkdirSync, renameSync, rmSync, realpathSync, statSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { errorText } from "../com/errors.ts";
import { loadConfig, purgeConfigAndState } from "../config/config.ts";
import {
  DAEMON_TASK,
  OUTLOOK_TASK,
  type TaskInfo,
  currentUser,
  findOutlookExe,
  hiddenLauncherScript,
  hiddenLauncherTaskSpec,
  outlookTaskSpec,
} from "../platform/scheduler.ts";
import { compareVersions } from "../util/version.ts";
import { VERSION } from "../version.ts";
import type { CommandContext } from "./context.ts";

const WINDOWS_ONLY = "this command is only available on Windows";

const samePath = (a: string, b: string): boolean => {
  try {
    return realpathSync(a).toLowerCase() === realpathSync(b).toLowerCase();
  } catch {
    return a.toLowerCase() === b.toLowerCase();
  }
};

function copyProgram(source: string, destination: string): void {
  mkdirSync(dirname(destination), { recursive: true });
  try {
    copyFileSync(source, destination);
  } catch (e) {
    // A running exe cannot be overwritten. Rename it aside, then copy.
    const aside = `${destination}.old-${process.pid}`;
    try {
      renameSync(destination, aside);
      copyFileSync(source, destination);
      rmSync(aside, { force: true });
    } catch {
      throw e;
    }
  }
}

export interface InstallOptions {
  noStart?: boolean;
}

export function runInstall(ctx: CommandContext, options: InstallOptions = {}): number {
  if (ctx.platform !== "win32") return fail(ctx, `${WINDOWS_ONLY} (${ctx.platform})`);
  if (!ctx.compiled) return fail(ctx, "install copies the built executable; run it from msoutlook-mcp.exe, not from bun");
  const { paths } = ctx;

  try {
    mkdirSync(paths.installDir, { recursive: true });
    ctx.tasks.stop(DAEMON_TASK); // a running copy locks its exe
    if (!samePath(ctx.execPath, paths.installedExe)) {
      copyProgram(ctx.execPath, paths.installedExe);
      ctx.out(`Copied the program to ${paths.installedExe}`);
    }
    // Keep the exe a console program; wscript starts it with no visible window.
    mkdirSync(dirname(paths.hiddenLauncher), { recursive: true });
    writeFileSync(
      paths.hiddenLauncher,
      Buffer.concat([
        Buffer.from([0xff, 0xfe]),
        Buffer.from(hiddenLauncherScript(paths.installedExe, "serve"), "utf16le"),
      ]),
    );
    const user = currentUser();
    ctx.tasks.create(DAEMON_TASK, hiddenLauncherTaskSpec(paths.hiddenLauncher, user, paths.installDir));
    ctx.out(`Registered the logon task "${DAEMON_TASK}" for ${user} (interactive, least privilege, no admin, no console window).`);
  } catch (e) {
    return fail(ctx, errorText(e));
  }

  // Outlook's own task is part of setup, but install keeps it consistent.
  try {
    const config = loadConfig(paths);
    if (config.start_outlook_at_logon) {
      const exe = findOutlookExe();
      if (exe) {
        ctx.tasks.create(OUTLOOK_TASK, outlookTaskSpec(exe, currentUser()));
        ctx.out(`Registered the Outlook logon task "${OUTLOOK_TASK}" (starts minimized).`);
      } else {
        ctx.out("Note: OUTLOOK.EXE was not found in App Paths; no Outlook logon task registered.");
      }
    }
  } catch {
    /* no config yet; setup will handle it */
  }

  if (options.noStart) {
    ctx.out("Not started (--no-start). Start it with: msoutlook-mcp install (without --no-start)");
    return 0;
  }
  try {
    ctx.tasks.start(DAEMON_TASK);
  } catch (e) {
    return fail(ctx, `installed, but the task did not start: ${errorText(e)}`);
  }
  ctx.out("Started. The daemon runs now and again at every sign-in.");
  return 0;
}

export interface UninstallOptions {
  purge?: boolean;
}

export function runUninstall(ctx: CommandContext, options: UninstallOptions = {}): number {
  if (ctx.platform !== "win32") return fail(ctx, `${WINDOWS_ONLY} (${ctx.platform})`);
  const { paths } = ctx;
  try {
    ctx.tasks.stop(DAEMON_TASK);
    ctx.tasks.stop(OUTLOOK_TASK);
    if (ctx.tasks.installed(DAEMON_TASK)) {
      ctx.tasks.delete(DAEMON_TASK);
      ctx.out(`Removed the logon task "${DAEMON_TASK}".`);
    } else {
      ctx.out("The daemon logon task was not installed.");
    }
    if (ctx.tasks.installed(OUTLOOK_TASK)) {
      ctx.tasks.delete(OUTLOOK_TASK);
      ctx.out(`Removed the Outlook logon task "${OUTLOOK_TASK}".`);
    }
    rmSync(paths.hiddenLauncher, { force: true });
    if (existsSync(paths.installedExe) && !samePath(ctx.execPath, paths.installedExe)) {
      const aside = `${paths.installedExe}.removed-${process.pid}`;
      try {
        renameSync(paths.installedExe, aside);
        rmSync(aside, { force: true });
        ctx.out(`Removed ${paths.installedExe}.`);
      } catch {
        ctx.out(`Could not remove ${paths.installedExe}; it is still running.`);
      }
    }
  } catch (e) {
    return fail(ctx, errorText(e));
  }

  if (options.purge) {
    try {
      purgeConfigAndState(paths);
      rmSync(paths.logDir, { recursive: true, force: true });
      ctx.out("Removed the config, token, state and logs (--purge).");
    } catch (e) {
      return fail(ctx, `could not purge local data: ${errorText(e)}`);
    }
  } else {
    ctx.out("Config and token kept. Use uninstall --purge to remove them too.");
  }
  return 0;
}

function installedVersion(ctx: CommandContext): string {
  if (!existsSync(ctx.paths.installedExe)) return "";
  const result = ctx.run([ctx.paths.installedExe, "version"]);
  return result.code === 0 ? result.out.trim() : "";
}

function taskInfoText(info: TaskInfo): string {
  if (!info.installed) return "not installed";
  const state = info.enabled ? "enabled" : "disabled";
  return `${state}, last result ${info.lastResult || "?"}, last run ${info.lastRunTime || "?"}`;
}

export { taskInfoText };

export interface UpgradeOptions {
  force?: boolean;
}

/**
 * The user runs the new exe as `<new>.exe upgrade`. No network access.
 * Stop the task, move the old exe aside, copy this exe in, run `version` on the
 * new copy, then restart. A failed self-check restores the old exe.
 */
export function runUpgrade(ctx: CommandContext, options: UpgradeOptions = {}): number {
  if (ctx.platform !== "win32") return fail(ctx, `upgrade replaces the Windows executable; not supported on ${ctx.platform}`);
  if (!ctx.compiled) return fail(ctx, "upgrade copies the built executable; run it from msoutlook-mcp.exe, not from bun");

  const { paths } = ctx;
  const newVersion = VERSION;
  const oldVersion = installedVersion(ctx);

  if (oldVersion && compareVersions(newVersion, oldVersion) < 0 && !options.force) {
    return fail(
      ctx,
      `refusing to downgrade from ${oldVersion} to ${newVersion}; pass --force if that is intended`,
    );
  }
  if (existsSync(paths.installedExe) && samePath(ctx.execPath, paths.installedExe)) {
    ctx.out(`Already running the installed copy (${newVersion}).`);
    return 0;
  }

  const wasInstalled = ctx.tasks.installed(DAEMON_TASK);
  const wasEnabled = wasInstalled ? ctx.tasks.query(DAEMON_TASK).enabled : false;
  const backup = `${paths.installedExe}.old-${process.pid}`;
  let moved = false;

  try {
    if (wasInstalled) ctx.tasks.stop(DAEMON_TASK);
    if (existsSync(paths.installedExe)) {
      renameSync(paths.installedExe, backup);
      moved = true;
    }
    copyProgram(ctx.execPath, paths.installedExe);
  } catch (e) {
    if (moved) {
      try {
        renameSync(backup, paths.installedExe);
      } catch {
        /* nothing else to do */
      }
    }
    if (wasEnabled) try {
      ctx.tasks.start(DAEMON_TASK);
    } catch {
      /* reported below */
    }
    return fail(ctx, `could not replace ${paths.installedExe}: ${errorText(e)}`);
  }

  const check = ctx.run([paths.installedExe, "version"]);
  const reported = check.out.trim();
  if (check.code !== 0 || reported !== newVersion) {
    try {
      rmSync(paths.installedExe, { force: true });
      if (moved) renameSync(backup, paths.installedExe);
    } catch {
      return fail(ctx, `the new copy failed its self-check and the old copy could not be restored. Backup: ${backup}`);
    }
    if (wasEnabled) {
      try {
        ctx.tasks.start(DAEMON_TASK);
      } catch {
        /* reported in the next line */
      }
    }
    return fail(
      ctx,
      `the new copy reported "${reported || "(nothing)"}" instead of ${newVersion}; the old copy was restored`,
    );
  }

  if (wasEnabled) {
    try {
      ctx.tasks.start(DAEMON_TASK);
    } catch (e) {
      return fail(ctx, `updated to ${newVersion}, but the task did not start: ${errorText(e)}`);
    }
  }
  if (moved) {
    try {
      rmSync(backup, { force: true });
    } catch {
      ctx.out(`Could not remove the old copy at ${backup}; delete it later.`);
    }
  }
  ctx.out(`Upgraded from ${oldVersion || "unknown"} to ${newVersion}.`);
  ctx.out("Check with: msoutlook-mcp doctor");
  return 0;
}

function fail(ctx: CommandContext, message: string): number {
  ctx.err(`error: ${message}`);
  return 1;
}

/** Used by doctor and tests without importing the executor twice. */
export function logFileInfo(ctx: CommandContext): { path: string; bytes: number; exists: boolean } {
  try {
    return { path: ctx.paths.logFile, bytes: statSync(ctx.paths.logFile).size, exists: true };
  } catch {
    return { path: ctx.paths.logFile, bytes: 0, exists: false };
  }
}

