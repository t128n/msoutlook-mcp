import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../src/com/fake.ts";
import { loadConfig, readToken, writeToken } from "../src/config/config.ts";
import { appPaths } from "../src/config/paths.ts";
import { MINIMIZE_HINT, mcpSnippet, runSetup } from "../src/commands/setup.ts";
import type { TaskInfo, TaskManager, TaskSpec } from "../src/platform/scheduler.ts";
import { OUTLOOK_TASK } from "../src/platform/scheduler.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

class FakeTasks implements TaskManager {
  created: { name: string; spec: TaskSpec }[] = [];
  installed(): boolean {
    return false;
  }
  query(): TaskInfo {
    return { installed: false, enabled: false, command: "", arguments: "", lastResult: "", lastRunTime: "", status: "" };
  }
  create(name: string, spec: TaskSpec): void {
    this.created.push({ name, spec });
  }
  delete(): void {}
  start(): void {}
  stop(): void {}
}

function temp(platform: NodeJS.Platform) {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-setup-"));
  dirs.push(dir);
  return { dir, paths: appPaths({ MSOUTLOOK_MCP_HOME: dir }, platform) };
}

function deps(platform: NodeJS.Platform, dir: string, tasks = new FakeTasks()) {
  const output: string[] = [];
  return {
    tasks,
    output,
    value: {
      paths: appPaths({ MSOUTLOOK_MCP_HOME: dir }, platform),
      platform,
      bridge: new FakeComBridge(),
      tasks,
      out: (text: string) => output.push(text),
      ask: async (_question: string, defaultValue: string) => defaultValue,
      portFree: async () => true,
      findOutlookExe: () => "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
      newOutlookDefault: () => false,
      currentUser: () => "EXAMPLE\\user",
    },
  };
}

describe("setup", () => {
  test("non-interactive setup writes config and a private token", async () => {
    const { dir, paths } = temp("darwin");
    const d = deps("darwin", dir);
    const code = await runSetup(
      {
        port: 9123,
        allowedRecipients: ["example.com", "*@contractor.example.net"],
        allowSend: false,
        startOutlookAtLogon: false,
        autoLaunchOutlook: false,
        interactive: false,
      },
      d.value,
    );
    expect(code).toBe(0);
    const config = loadConfig(paths);
    expect(config.port).toBe(9123);
    expect(config.allowed_recipients).toEqual(["example.com", "*@contractor.example.net"]);
    expect(config.allow_send).toBe(false);
    expect(config.auto_launch_outlook).toBe(false);
    const token = readToken(paths);
    expect(token.length).toBeGreaterThan(30);
    if (process.platform !== "win32") expect(statSync(paths.tokenFile).mode & 0o077).toBe(0);
    const printed = d.output.join("\n");
    expect(printed).toContain(`http://127.0.0.1:9123/mcp`);
    expect(printed).toContain(token);
    expect(printed).toContain("Minimize Outlook");
  });

  test("prints a usable MCP snippet", () => {
    const snippet = JSON.parse(mcpSnippet(8787, "secret")) as {
      mcpServers: { outlook: { type: string; url: string; headers: { Authorization: string } } };
    };
    expect(snippet.mcpServers.outlook.type).toBe("remote");
    expect(snippet.mcpServers.outlook.url).toBe("http://127.0.0.1:8787/mcp");
    expect(snippet.mcpServers.outlook.headers.Authorization).toBe("Bearer secret");
  });

  test("registers the Outlook logon task on Windows when asked", async () => {
    const { dir } = temp("win32");
    const d = deps("win32", dir);
    const code = await runSetup(
      {
        port: 9124,
        allowedRecipients: ["example.com"],
        allowSend: true,
        startOutlookAtLogon: true,
        autoLaunchOutlook: false,
        interactive: false,
      },
      d.value,
    );
    expect(code).toBe(0);
    expect(d.tasks.created).toHaveLength(1);
    expect(d.tasks.created[0]!.name).toBe(OUTLOOK_TASK);
    expect(d.tasks.created[0]!.spec.args).toContain('/min "');
    expect(d.output.join("\n")).toContain("Hide When Minimized");
  });

  test("no Outlook task on non-Windows", async () => {
    const { dir } = temp("darwin");
    const d = deps("darwin", dir);
    const code = await runSetup(
      {
        startOutlookAtLogon: true,
        interactive: false,
      },
      d.value,
    );
    expect(code).toBe(0);
    expect(d.tasks.created).toHaveLength(0);
  });

  test("an invalid allowlist stops setup before writing config", async () => {
    const { dir, paths } = temp("darwin");
    const d = deps("darwin", dir);
    await expect(
      runSetup(
        {
          allowedRecipients: ["/[a-/"],
          interactive: false,
        },
        d.value,
      ),
    ).rejects.toThrow(/invalid regular expression/);
    expect(() => loadConfig(paths)).toThrow();
  });

  test("keeps an existing token unless rotation is requested", async () => {
    const { dir, paths } = temp("darwin");
    const first = writeToken(paths);
    const d = deps("darwin", dir);
    await runSetup({ interactive: false, startOutlookAtLogon: false }, d.value);
    expect(readToken(paths)).toBe(first);

    const d2 = deps("darwin", dir);
    await runSetup({ interactive: false, rotateToken: true, startOutlookAtLogon: false }, d2.value);
    expect(readToken(paths)).not.toBe(first);
  });

  test("the minimize hint names the real Outlook menu item", () => {
    expect(MINIMIZE_HINT).toContain("Minimize Outlook");
    expect(MINIMIZE_HINT).toContain("Hide When Minimized");
    expect(MINIMIZE_HINT).toContain("do not close");
  });

  test("warns when Outlook is not reachable but still writes config", async () => {
    const { dir, paths } = temp("darwin");
    const d = deps("darwin", dir);
    (d.value.bridge as FakeComBridge).closeOutlook();
    const code = await runSetup({ interactive: false, startOutlookAtLogon: false }, d.value);
    expect(code).toBe(0);
    expect(() => loadConfig(paths)).not.toThrow();
    expect(d.output.join("\n")).toContain("Not reachable yet");
  });

  test("the config file contains no token", async () => {
    const { dir, paths } = temp("darwin");
    const d = deps("darwin", dir);
    await runSetup({ interactive: false, startOutlookAtLogon: false }, d.value);
    const token = readToken(paths);
    expect(readFileSync(paths.configFile, "utf8")).not.toContain(token);
  });
});
