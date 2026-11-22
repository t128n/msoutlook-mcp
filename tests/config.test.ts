import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  DEFAULT_CONFIG,
  ConfigError,
  loadConfig,
  readToken,
  saveConfig,
  validateConfig,
  writeToken,
} from "../src/config/config.ts";
import { appPaths } from "../src/config/paths.ts";

const dirs: string[] = [];
function tempPaths() {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-config-"));
  dirs.push(dir);
  return appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin");
}
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

describe("config", () => {
  test("defaults are safe: sending off, auto-launch off", () => {
    expect(DEFAULT_CONFIG.allow_send).toBe(false);
    expect(DEFAULT_CONFIG.auto_launch_outlook).toBe(false);
    expect(DEFAULT_CONFIG.allowed_recipients).toEqual([]);
    expect(DEFAULT_CONFIG.poll_interval_ms).toBe(30_000);
  });

  test("round-trips through disk", () => {
    const paths = tempPaths();
    const config = { ...DEFAULT_CONFIG, port: 9123, allow_send: true, allowed_recipients: ["example.com"] };
    saveConfig(paths, config);
    expect(loadConfig(paths)).toEqual(config);
  });

  test("missing config explains setup", () => {
    const paths = tempPaths();
    expect(() => loadConfig(paths)).toThrow(ConfigError);
    try {
      loadConfig(paths);
    } catch (e) {
      expect((e as ConfigError).hint).toContain("setup");
    }
  });

  test("an invalid regex stops the config from loading", () => {
    const paths = tempPaths();
    writeFileSync(paths.configFile, JSON.stringify({ ...DEFAULT_CONFIG, allowed_recipients: ["/[a-/"] }));
    expect(() => loadConfig(paths)).toThrow(/invalid regular expression/);
  });

  test("validation rejects impossible values", () => {
    expect(() => validateConfig({ port: 0 })).toThrow(ConfigError);
    expect(() => validateConfig({ port: 70_000 })).toThrow(ConfigError);
    expect(() => validateConfig({ allow_send: "yes" })).toThrow(ConfigError);
    expect(() => validateConfig({ poll_interval_ms: 10 })).toThrow(ConfigError);
    expect(() => validateConfig({ log_level: "loud" })).toThrow(ConfigError);
    expect(() => validateConfig([])).toThrow(ConfigError);
  });

  test("unknown keys are ignored and defaults fill in", () => {
    const config = validateConfig({ port: 9000, something_new: true });
    expect(config.port).toBe(9000);
    expect(config.allow_send).toBe(false);
    expect(config.start_outlook_at_logon).toBe(true);
  });

  test("shared mailbox names are lowercased", () => {
    const config = validateConfig({ shared_mailboxes: ["Team@Example.com"] });
    expect(config.shared_mailboxes).toEqual(["team@example.com"]);
  });

  test("token is 0600 and not part of config", () => {
    const paths = tempPaths();
    const token = writeToken(paths);
    expect(token.length).toBeGreaterThan(30);
    expect(readToken(paths)).toBe(token);
    if (process.platform !== "win32") {
      expect(statSync(paths.tokenFile).mode & 0o077).toBe(0);
    }
    const config = { ...DEFAULT_CONFIG };
    saveConfig(paths, config);
    expect(readFileSync(paths.configFile, "utf8")).not.toContain(token);
  });

  test("the token file is separate and empty tokens are refused", () => {
    const paths = tempPaths();
    writeFileSync(paths.tokenFile, "\n");
    expect(() => readToken(paths)).toThrow(/empty/);
  });
});
