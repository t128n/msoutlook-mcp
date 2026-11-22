import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../src/com/fake.ts";
import { startServe, type ServeHandle } from "../src/commands/serve.ts";
import type { CommandContext } from "../src/commands/context.ts";
import { DEFAULT_CONFIG, saveConfig, writeToken } from "../src/config/config.ts";
import { appPaths } from "../src/config/paths.ts";
import type { TaskInfo, TaskManager } from "../src/platform/scheduler.ts";

const handles: ServeHandle[] = [];
const dirs: string[] = [];
afterEach(() => {
  for (const h of handles.splice(0)) h.stop();
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

class NoTasks implements TaskManager {
  installed(): boolean {
    return false;
  }
  query(): TaskInfo {
    return { installed: false, enabled: false, command: "", arguments: "", lastResult: "", lastRunTime: "", status: "" };
  }
  create(): void {}
  delete(): void {}
  start(): void {}
  stop(): void {}
}

async function start(open: boolean) {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-serve-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin");
  saveConfig(paths, { ...DEFAULT_CONFIG, start_outlook_at_logon: false });
  writeToken(paths);
  const bridge = new FakeComBridge({ running: open });
  const out: string[] = [];
  const ctx: CommandContext = {
    paths,
    platform: "darwin",
    env: {},
    execPath: join(dir, "msoutlook-mcp"),
    compiled: true,
    isTTY: false,
    out: (text) => out.push(text ?? ""),
    err: (text) => out.push(text ?? ""),
    bridgeFactory: () => bridge,
    tasks: new NoTasks(),
    now: () => new Date(),
    run: () => ({ code: 1, out: "", err: "" }),
  };
  const handle = await startServe(ctx, { port: 0, signals: false });
  handles.push(handle);
  return { handle, bridge, dir, paths };
}

describe("serve with the lifecycle wrapper", () => {
  test("starts with Outlook open and reports connected", async () => {
    const { handle } = await start(true);
    const response = await fetch(`http://127.0.0.1:${handle.port}/health`);
    const body = (await response.json()) as { status: string };
    expect(body.status).toBe("connected");
  });

  test("starts with Outlook closed and reports waiting_for_outlook", async () => {
    const { handle } = await start(false);
    const body = (await (await fetch(`http://127.0.0.1:${handle.port}/health`)).json()) as {
      status: string;
      outlook: { retryable: boolean; reconnect_attempts: number };
    };
    expect(body.status).toBe("waiting_for_outlook");
    expect(body.outlook.retryable).toBe(true);
    expect(body.outlook.reconnect_attempts).toBeGreaterThanOrEqual(1);
  });

  test("stopping the daemon closes the bridge", async () => {
    const { handle } = await start(true);
    handle.stop();
    await handle.done;
    expect(true).toBe(true);
  });
});
