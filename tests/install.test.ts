import { afterEach, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { CommandContext } from "../src/commands/context.ts";
import { runInstall, runUninstall } from "../src/commands/lifecycle.ts";
import { DEFAULT_CONFIG, saveConfig } from "../src/config/config.ts";
import { appPaths } from "../src/config/paths.ts";
import type { TaskInfo, TaskManager, TaskSpec } from "../src/platform/scheduler.ts";
import { DAEMON_TASK, OUTLOOK_TASK } from "../src/platform/scheduler.ts";

const dirs: string[] = [];
process.env.USERDOMAIN = "EXAMPLE";
process.env.USERNAME = "user";
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

class FakeTasks implements TaskManager {
  installedFlag = true;
  created: { name: string; spec: TaskSpec }[] = [];
  deleted: string[] = [];
  calls: string[] = [];
  installed(): boolean {
    return this.installedFlag;
  }
  query(): TaskInfo {
    return { installed: true, enabled: true, command: "wscript.exe", arguments: "", lastResult: "0", lastRunTime: "today", status: "Ready" };
  }
  create(name: string, spec: TaskSpec): void {
    this.created.push({ name, spec });
  }
  delete(name: string): void {
    this.deleted.push(name);
  }
  start(name: string): void {
    this.calls.push(`start:${name}`);
  }
  stop(name: string): void {
    this.calls.push(`stop:${name}`);
  }
}

function setup() {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-install-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "win32");
  const source = join(dir, "downloaded.exe");
  writeFileSync(source, "the-new-exe");
  saveConfig(paths, { ...DEFAULT_CONFIG, start_outlook_at_logon: false });
  const tasks = new FakeTasks();
  const out: string[] = [];
  const ctx: CommandContext = {
    paths,
    platform: "win32",
    env: {},
    execPath: source,
    compiled: true,
    isTTY: false,
    out: (text) => out.push(text ?? ""),
    err: (text) => out.push(`ERR ${text ?? ""}`),
    bridgeFactory: () => {
      throw new Error("not used");
    },
    tasks,
    now: () => new Date(),
    run: () => ({ code: 0, out: "", err: "" }),
  };
  return { dir, paths, source, tasks, out, ctx };
}

describe("install and uninstall", () => {
  test("copies the exe, writes a hidden launcher and starts the task", () => {
    const { paths, tasks, ctx } = setup();
    expect(runInstall(ctx)).toBe(0);
    expect(readFileSync(paths.installedExe, "utf8")).toBe("the-new-exe");
    expect(existsSync(paths.hiddenLauncher)).toBe(true);
    expect(readFileSync(paths.hiddenLauncher, "utf16le")).toContain(paths.installedExe);
    const task = tasks.created.find((t) => t.name === DAEMON_TASK);
    expect(task?.spec.exe).toBe("wscript.exe");
    expect(task?.spec.args).toContain(paths.hiddenLauncher);
    expect(tasks.calls).toContain(`stop:${DAEMON_TASK}`);
    expect(tasks.calls).toContain(`start:${DAEMON_TASK}`);
  });

  test("--no-start registers without starting", () => {
    const { tasks, ctx } = setup();
    expect(runInstall(ctx, { noStart: true })).toBe(0);
    expect(tasks.calls).not.toContain(`start:${DAEMON_TASK}`);
  });

  test("refuses to install from an interpreted process", () => {
    const { ctx, out } = setup();
    expect(runInstall({ ...ctx, compiled: false })).toBe(1);
    expect(out.join("\n")).toContain("built executable");
  });

  test("uninstall removes both tasks and the exe but keeps config", () => {
    const { paths, tasks, ctx } = setup();
    runInstall(ctx);
    const code = runUninstall(ctx);
    expect(code).toBe(0);
    expect(tasks.deleted).toContain(DAEMON_TASK);
    expect(tasks.deleted).toContain(OUTLOOK_TASK);
    expect(existsSync(paths.installedExe)).toBe(false);
    expect(existsSync(paths.hiddenLauncher)).toBe(false);
    expect(existsSync(paths.configFile)).toBe(true);
  });

  test("uninstall --purge removes config and token", () => {
    const { paths, ctx } = setup();
    runInstall(ctx);
    expect(runUninstall(ctx, { purge: true })).toBe(0);
    expect(existsSync(paths.configFile)).toBe(false);
  });
});
