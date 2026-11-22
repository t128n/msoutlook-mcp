import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../src/com/fake.ts";
import { main, parseFlags, UsageError } from "../src/cli.ts";
import { loadConfig, readToken, writeToken } from "../src/config/config.ts";
import { appPaths } from "../src/config/paths.ts";
import type { TaskInfo, TaskManager } from "../src/platform/scheduler.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

class FakeTasks implements TaskManager {
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

function harness(platform: NodeJS.Platform = "darwin") {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-cli-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, platform);
  const out: string[] = [];
  const err: string[] = [];
  const ctx = {
    paths,
    platform,
    env: { MSOUTLOOK_MCP_HOME: dir },
    execPath: join(dir, "msoutlook-mcp"),
    compiled: true,
    isTTY: false,
    out: (text?: string) => out.push(text ?? ""),
    err: (text?: string) => err.push(text ?? ""),
    bridgeFactory: () => new FakeComBridge(),
    tasks: new FakeTasks(),
    now: () => new Date(),
    run: () => ({ code: 1, out: "", err: "" }),
  };
  return { dir, paths, out, err, ctx };
}

describe("flag parsing", () => {
  test("booleans, negations, inline strings and positionals", () => {
    const parsed = parseFlags(
      ["--allow-send", "--no-outlook-logon", "--port=9123", "--allow", "a@example.com", "extra"],
      {
        "allow-send": { type: "boolean", description: "" },
        "outlook-logon": { type: "boolean", description: "" },
        port: { type: "string", description: "", arg: "port" },
        allow: { type: "string", description: "", arg: "list" },
      },
    );
    expect(parsed.values["allow-send"]).toBe(true);
    expect(parsed.values["outlook-logon"]).toBe(false);
    expect(parsed.values.port).toBe("9123");
    expect(parsed.values.allow).toBe("a@example.com");
    expect(parsed.positionals).toEqual(["extra"]);
  });

  test("unknown flags and missing values are usage errors", () => {
    expect(() => parseFlags(["--nope"], {})).toThrow(UsageError);
    expect(() => parseFlags(["--port"], { port: { type: "string", description: "" } })).toThrow(UsageError);
  });

  test("--help is recognised anywhere", () => {
    expect(parseFlags(["setup", "--help"], {}).help).toBe(true);
  });
});

describe("CLI", () => {
  test("version prints the version", async () => {
    const { out, ctx } = harness();
    expect(await main(["version"], ctx)).toBe(0);
    expect(out.join("\n")).toMatch(/^\d+\.\d+\.\d+/);
  });

  test("--help lists every command", async () => {
    const { out, ctx } = harness();
    expect(await main(["--help"], ctx)).toBe(0);
    const text = out.join("\n");
    for (const command of ["setup", "doctor", "install", "uninstall", "upgrade", "serve", "version"]) {
      expect(text).toContain(command);
    }
  });

  test("each command has --help", async () => {
    for (const command of ["setup", "doctor", "install", "uninstall", "upgrade", "serve", "config"]) {
      const { out, ctx } = harness();
      expect(await main([command, "--help"], ctx)).toBe(0);
      expect(out.join("\n")).toContain("Usage:");
    }
  });

  test("setup writes config and keeps the redesign's defaults", async () => {
    const { paths, out, ctx } = harness();
    const code = await main(
      [
        "setup",
        "--non-interactive",
        "--port",
        "19123",
        "--allow",
        "example.com,*@contractor.example.net",
        "--no-outlook-logon",
      ],
      ctx,
    );
    expect(code).toBe(0);
    const config = loadConfig(paths);
    expect(config.port).toBe(19123);
    expect(config.allow_send).toBe(false);
    expect(config.allowed_recipients).toEqual(["example.com", "*@contractor.example.net"]);
    expect(out.join("\n")).toContain("http://127.0.0.1:19123/mcp");
  });

  test("config set flips sending without a restart", async () => {
    const { paths, out, ctx } = harness();
    await main(["setup", "--non-interactive", "--port", "19124", "--no-outlook-logon"], ctx);
    expect(await main(["config", "set", "allow-send", "on"], ctx)).toBe(0);
    expect(loadConfig(paths).allow_send).toBe(true);
    expect(out.join("\n")).toContain("read at send time");

    const showOut: string[] = [];
    expect(await main(["config", "show"], { ...ctx, out: (t) => showOut.push(t ?? "") })).toBe(0);
    expect(showOut.join("\n")).toContain('"allow_send": true');
    expect(showOut.join("\n")).not.toContain(readToken(paths));
  });

  test("config set rejects an invalid regex", async () => {
    const { paths, err, ctx } = harness();
    await main(["setup", "--non-interactive", "--no-outlook-logon"], ctx);
    expect(await main(["config", "set", "allowed-recipients", "/[a-/"], ctx)).toBe(1);
    expect(err.join("\n")).toContain("invalid regular expression");
    expect(loadConfig(paths).allowed_recipients).toEqual([]);
  });

  test("doctor prints one line per check and exits non-zero off Windows", async () => {
    const { paths, out, ctx } = harness();
    await main(["setup", "--non-interactive", "--no-outlook-logon"], ctx);
    const code = await main(["doctor"], ctx);
    expect(code).toBe(1);
    const text = out.join("\n");
    expect(text).toContain("Windows:");
    expect(text).toContain("Config:");
    expect(text).toContain("Guardrails:");
    expect(text).toContain("Keep Outlook:");
    expect(text).toContain("checks,");
  });

  test("install is refused off Windows", async () => {
    const { err, ctx } = harness();
    expect(await main(["install"], ctx)).toBe(1);
    expect(err.join("\n")).toContain("only available on Windows");
  });

  test("unknown commands exit 2", async () => {
    const { err, ctx } = harness();
    expect(await main(["wat"], ctx)).toBe(2);
    expect(err.join("\n")).toContain("unknown command");
  });

  test("a missing config explains setup", async () => {
    const { err, ctx } = harness();
    expect(await main(["serve"], ctx)).toBe(1);
    expect(err.join("\n")).toContain("setup");
  });

  test("serve --port rejects nonsense", async () => {
    const { err, ctx } = harness();
    expect(await main(["serve", "--port", "nope"], ctx)).toBe(2);
    expect(err.join("\n")).toContain("--port");
  });

  test("setup rotates a token only when asked", async () => {
    const { paths, ctx } = harness();
    await main(["setup", "--non-interactive", "--no-outlook-logon"], ctx);
    const first = readToken(paths);
    await main(["setup", "--non-interactive", "--no-outlook-logon"], ctx);
    expect(readToken(paths)).toBe(first);
    writeToken(paths);
    await main(["setup", "--non-interactive", "--no-outlook-logon", "--rotate-token"], ctx);
    expect(readToken(paths)).not.toBe(first);
  });
});
