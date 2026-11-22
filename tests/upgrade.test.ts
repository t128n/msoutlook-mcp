import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { CommandContext } from "../src/commands/context.ts";
import { runUpgrade } from "../src/commands/lifecycle.ts";
import { appPaths } from "../src/config/paths.ts";
import type { TaskInfo, TaskManager } from "../src/platform/scheduler.ts";
import { DAEMON_TASK } from "../src/platform/scheduler.ts";
import { VERSION } from "../src/version.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

class FakeTasks implements TaskManager {
  installedFlag = true;
  enabled = true;
  calls: string[] = [];
  installed(): boolean {
    return this.installedFlag;
  }
  query(): TaskInfo {
    return {
      installed: this.installedFlag,
      enabled: this.enabled,
      command: "x",
      arguments: "serve",
      lastResult: "0",
      lastRunTime: "today",
      status: "Ready",
    };
  }
  create(): void {}
  delete(): void {
    this.calls.push("delete");
  }
  start(): void {
    this.calls.push("start");
  }
  stop(): void {
    this.calls.push("stop");
  }
}

function setup(options: { installedContent?: string; newContent?: string; installedReports?: string } = {}) {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-upgrade-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "win32");
  const newExe = join(dir, "new.exe");
  writeFileSync(newExe, options.newContent ?? "new");
  writeFileSync(paths.installedExe, options.installedContent ?? "old");
  const tasks = new FakeTasks();
  const output: string[] = [];
  const installedReports = options.installedReports ?? "0.0.9";
  const ctx: CommandContext = {
    paths,
    platform: "win32",
    env: {},
    execPath: newExe,
    compiled: true,
    isTTY: false,
    out: (text) => output.push(text ?? ""),
    err: (text) => output.push(`ERR ${text ?? ""}`),
    bridgeFactory: () => {
      throw new Error("not used");
    },
    tasks,
    now: () => new Date(),
    run: (cmd) => {
      if (cmd[0] === paths.installedExe) {
        const content = readFileSync(paths.installedExe, "utf8");
        if (content === "old") return { code: 0, out: installedReports, err: "" };
        if (content === "new") return { code: 0, out: VERSION, err: "" };
        return { code: 0, out: "0.0.0", err: "" };
      }
      return { code: 1, out: "", err: "unexpected command" };
    },
  };
  return { dir, paths, newExe, tasks, output, ctx };
}

describe("upgrade", () => {
  test("replaces the installed exe, self-checks, restarts and prints both versions", () => {
    const { paths, tasks, output, ctx } = setup();
    expect(runUpgrade(ctx)).toBe(0);
    expect(readFileSync(paths.installedExe, "utf8")).toBe("new");
    expect(tasks.calls).toEqual(["stop", "start"]);
    expect(output.join("\n")).toContain(`Upgraded from 0.0.9 to ${VERSION}`);
  });

  test("refuses a downgrade without --force", () => {
    const { paths, tasks, output, ctx } = setup({ installedReports: "99.0.0" });
    expect(runUpgrade(ctx)).toBe(1);
    expect(readFileSync(paths.installedExe, "utf8")).toBe("old");
    expect(tasks.calls).toEqual([]);
    expect(output.join("\n")).toContain("refusing to downgrade");
  });

  test("--force allows a downgrade", () => {
    const { paths, tasks, ctx } = setup({ installedReports: "99.0.0" });
    expect(runUpgrade(ctx, { force: true })).toBe(0);
    expect(readFileSync(paths.installedExe, "utf8")).toBe("new");
    expect(tasks.calls).toEqual(["stop", "start"]);
  });

  test("restores the old exe when the new copy fails its self-check", () => {
    const { paths, tasks, output, ctx } = setup({ newContent: "bad" });
    expect(runUpgrade(ctx)).toBe(1);
    expect(readFileSync(paths.installedExe, "utf8")).toBe("old");
    expect(tasks.calls).toEqual(["stop", "start"]);
    expect(output.join("\n")).toContain("the old copy was restored");
  });

  test("refuses to run from an interpreted bun process", () => {
    const { ctx } = setup();
    expect(runUpgrade({ ...ctx, compiled: false })).toBe(1);
  });

  test("is Windows only", () => {
    const { ctx } = setup();
    expect(runUpgrade({ ...ctx, platform: "darwin" })).toBe(1);
  });

  test("the task name is the daemon task", () => {
    expect(DAEMON_TASK).toBe("MSOutlook MCP");
  });
});
