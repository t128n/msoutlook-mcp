import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../src/com/fake.ts";
import { DEFAULT_CONFIG, type AppConfig } from "../src/config/config.ts";
import { Guardrails } from "../src/config/guardrails.ts";
import { appPaths } from "../src/config/paths.ts";
import { HealthMonitor } from "../src/server/health.ts";
import { createTools } from "../src/tools/index.ts";
import type { ToolResponse } from "../src/tools/registry.ts";
import { Logger } from "../src/util/logger.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

function setup(config: Partial<AppConfig> = {}) {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-rules-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin");
  const bridge = new FakeComBridge();
  const guardrails = new Guardrails({ ...DEFAULT_CONFIG, allowed_recipients: ["example.com"], ...config });
  const health = new HealthMonitor({ bridge, version: "0.1.0", cacheMs: 0, timeoutMs: 200 });
  const tools = createTools({
    bridge,
    guardrails,
    health,
    paths,
    version: "0.1.0",
    logger: new Logger(join(dir, "server.log"), "error"),
    now: () => new Date(),
  });
  const call = async (name: string, args: Record<string, unknown> = {}) =>
    (await tools.call(name, args)) as ToolResponse;
  return { bridge, tools, call };
}

describe("rules", () => {
  test("list_rules decodes existing rules", async () => {
    const { bridge, call } = setup();
    bridge.seedRule({ name: "Move invoices", conditions: ["Subject CONTAINS invoice"], actions: ["Move -> Invoices"] });
    const result = await call("list_rules");
    const data = result.data as { count: number; rules: { name: string; conditions: string[] }[] };
    expect(data.count).toBe(1);
    expect(data.rules[0]!.name).toBe("Move invoices");
    expect(data.rules[0]!.conditions).toEqual(["Subject CONTAINS invoice"]);
  });

  test("toggle_rule needs confirm and changes the rule", async () => {
    const { bridge, call } = setup();
    bridge.seedRule({ name: "Rule A", enabled: true });
    const noConfirm = await call("toggle_rule", { name: "Rule A", enabled: false });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    expect(bridge.state.rules[0]!.enabled).toBe(true);
    const result = await call("toggle_rule", { name: "Rule A", enabled: false, confirm: true });
    expect(result.isError).toBeUndefined();
    expect(bridge.state.rules[0]!.enabled).toBe(false);
  });

  test("create_rule requires a condition, an action and confirm", async () => {
    const { call } = setup();
    const noCondition = await call("create_rule", { name: "x", subject_contains: [], confirm: true });
    expect((noCondition.data as { error: string }).error).toBe("NO_CONDITION");
    const noAction = await call("create_rule", {
      name: "x",
      subject_contains: ["invoice"],
      confirm: true,
    });
    expect((noAction.data as { error: string }).error).toBe("NO_ACTION");
    const noConfirm = await call("create_rule", {
      name: "x",
      subject_contains: ["invoice"],
      move_to_folder: "archive",
    });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
  });

  test("create_rule checks forward addresses but not incoming senders", async () => {
    const { bridge, call } = setup();
    const blocked = await call("create_rule", {
      name: "x",
      from_addresses: ["outside@example.org"],
      subject_contains: ["invoice"],
      forward_to: ["outside@example.org"],
      confirm: true,
    });
    expect((blocked.data as { error: string }).error).toBe("ADDRESSES_NOT_ALLOWED");
    expect(bridge.state.rules).toHaveLength(0);

    const senderOutside = await call("create_rule", {
      name: "Senders",
      from_addresses: ["outside@example.org"],
      subject_contains: ["invoice"],
      move_to_folder: "archive",
      confirm: true,
    });
    expect(senderOutside.isError).toBeUndefined();
    expect(bridge.state.rules).toHaveLength(1);
  });

  test("a forward rule obeys the send switch", async () => {
    const off = setup();
    const blocked = await off.call("create_rule", {
      name: "Forward",
      subject_contains: ["invoice"],
      forward_to: ["one@example.com"],
      confirm: true,
    });
    expect((blocked.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    expect(off.bridge.state.rules).toHaveLength(0);

    const on = setup({ allow_send: true });
    const ok = await on.call("create_rule", {
      name: "Forward",
      subject_contains: ["invoice"],
      forward_to: ["one@example.com"],
      confirm: true,
    });
    expect(ok.isError).toBeUndefined();
    expect(on.bridge.state.rules).toHaveLength(1);
  });

  test("enabling a forward rule obeys the send switch", async () => {
    const { bridge, call } = setup();
    bridge.seedRule({ name: "Forward", enabled: false, actions: ["Forward -> [one@example.com]"] });
    const blocked = await call("toggle_rule", { name: "Forward", enabled: true, confirm: true });
    expect((blocked.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    expect(bridge.state.rules[0]!.enabled).toBe(false);
    const disabled = await call("toggle_rule", { name: "Forward", enabled: false, confirm: true });
    expect(disabled.isError).toBeUndefined();
  });

  test("create_rule saves a full rule and refuses a duplicate name", async () => {
    const { bridge, call } = setup({ allow_send: true });
    const result = await call("create_rule", {
      name: "Invoices",
      subject_contains: ["invoice", "bill"],
      body_or_subject_contains: ["payment"],
      has_attachment: true,
      importance_high: true,
      move_to_folder: "archive",
      copy_to_folder: "archive",
      assign_categories: ["Finance"],
      forward_to: ["one@example.com"],
      stop_processing: true,
      enabled: true,
      apply_now: true,
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    const data = result.data as { ok: boolean; apply_now_executed: boolean; summary: { conditions: string[]; actions: string[] } };
    expect(data.ok).toBe(true);
    expect(data.apply_now_executed).toBe(true);
    expect(data.summary.conditions).toContain("Subject CONTAINS invoice, bill");
    expect(data.summary.actions).toContain("Forward -> [one@example.com]");
    const duplicate = await call("create_rule", {
      name: "Invoices",
      subject_contains: ["x"],
      move_to_folder: "archive",
      confirm: true,
    });
    expect((duplicate.data as { error: string }).error).toBe("RULE_NAME_EXISTS");
  });

  test("delete_rule needs confirm and a real rule", async () => {
    const { bridge, call } = setup();
    bridge.seedRule({ name: "Rule A" });
    const noConfirm = await call("delete_rule", { name: "Rule A" });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    const missing = await call("delete_rule", { name: "Nope", confirm: true });
    expect((missing.data as { error: string }).error).toBe("RULE_NOT_FOUND");
    const result = await call("delete_rule", { name: "Rule A", confirm: true });
    expect(result.isError).toBeUndefined();
    expect(bridge.state.rules).toHaveLength(0);
  });
});

describe("contacts", () => {
  test("search_contacts returns GAL matches", async () => {
    const { bridge, call } = setup();
    bridge.seedContact({ name: "Ada Lovelace", smtp_address: "ada@example.com", department: "Research" });
    bridge.seedContact({ name: "Grace Hopper", smtp_address: "grace@example.com" });
    const result = await call("search_contacts", { query: "ada" });
    const data = result.data as { count: number; contacts: { name: string; smtp_address: string }[] };
    expect(data.count).toBe(1);
    expect(data.contacts[0]).toMatchObject({ name: "Ada Lovelace", smtp_address: "ada@example.com" });
  });

  test("search_contacts respects the limit", async () => {
    const { bridge, call } = setup();
    for (let i = 0; i < 5; i++) bridge.seedContact({ name: `Person ${i}` });
    const result = await call("search_contacts", { query: "Person", limit: 2 });
    expect((result.data as { count: number }).count).toBe(2);
  });

  test("get_contact_details returns address and manager data", async () => {
    const { bridge, call } = setup();
    bridge.seedContact({
      name: "Ada Lovelace",
      smtp_address: "ada@example.com",
      city: "London",
      manager: { name: "Charles Babbage", smtp_address: "charles@example.com" },
    });
    const result = await call("get_contact_details", { smtp_address: "ada@example.com" });
    const data = result.data as { name: string; city: string; manager: { name: string } };
    expect(data.name).toBe("Ada Lovelace");
    expect(data.city).toBe("London");
    expect(data.manager.name).toBe("Charles Babbage");
  });

  test("a missing contact returns CONTACT_NOT_FOUND", async () => {
    const { call } = setup();
    const result = await call("get_contact_details", { smtp_address: "nobody@example.com" });
    expect(result.isError).toBe(true);
    expect((result.data as { error: string }).error).toBe("CONTACT_NOT_FOUND");
  });

  test("the registry contains all 27 tools", async () => {
    const { tools } = setup();
    expect(tools.list()).toHaveLength(27);
    for (const name of [
      "search_contacts",
      "get_contact_details",
      "list_rules",
      "toggle_rule",
      "create_rule",
      "delete_rule",
    ]) {
      expect(tools.list().map((t) => t.name)).toContain(name);
    }
  });
});
