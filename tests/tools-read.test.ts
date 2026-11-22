import { afterEach, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
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

function setup(config: Partial<AppConfig> = {}, now?: Date) {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-tools-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin");
  const bridge = new FakeComBridge();
  const fullConfig: AppConfig = { ...DEFAULT_CONFIG, ...config };
  const guardrails = new Guardrails(fullConfig);
  const health = new HealthMonitor({ bridge, version: "0.1.0", cacheMs: 0, timeoutMs: 200 });
  const clock = now ?? new Date();
  const tools = createTools({
    bridge,
    guardrails,
    health,
    paths,
    version: "0.1.0",
    logger: new Logger(join(dir, "server.log"), "error"),
    now: () => clock,
  });
  const call = async (name: string, args: Record<string, unknown> = {}) =>
    (await tools.call(name, args)) as ToolResponse;
  return { dir, paths, bridge, guardrails, health, tools, call, clock };
}

describe("read tools", () => {
  test("list_folders returns the tree", async () => {
    const { call } = setup();
    const result = await call("list_folders", { max_depth: 2 });
    expect(result.isError).toBeUndefined();
    const folders = result.data as { name: string; children: unknown[] }[];
    expect(folders[0]!.name).toBe("Personal Mailbox");
    expect(folders[0]!.children.length).toBeGreaterThan(0);
  });

  test("list_mail returns summaries, truncated and coverage, never a body", async () => {
    const { bridge, call } = setup();
    const now = Date.now();
    bridge.seedMail({ subject: "newer", received_at: new Date(now - 1_000).toISOString(), body: "secret body" });
    bridge.seedMail({ subject: "older", received_at: new Date(now - 2_000).toISOString() });
    const result = await call("list_mail", { limit: 10 });
    const data = result.data as {
      count: number;
      truncated: boolean;
      coverage: { oldest_received: string; newest_received: string };
      items: { subject: string; body?: unknown }[];
    };
    expect(data.count).toBe(2);
    expect(data.truncated).toBe(false);
    expect(data.items[0]!.subject).toBe("newer");
    expect(data.items[0]!.body).toBeUndefined();
    expect(data.coverage.newest_received).toBe(new Date(now - 1_000).toISOString());
    expect(data.coverage.oldest_received).toBe(new Date(now - 2_000).toISOString());
  });

  test("list_mail marks a capped window as truncated", async () => {
    const { bridge, call } = setup();
    bridge.seedMail({ subject: "a", received_at: new Date(Date.now() - 1_000).toISOString() });
    bridge.seedMail({ subject: "b", received_at: new Date(Date.now() - 2_000).toISOString() });
    const result = await call("list_mail", { limit: 1 });
    const data = result.data as { count: number; truncated: boolean };
    expect(data.count).toBe(1);
    expect(data.truncated).toBe(true);
  });

  test("list_mail applies unread_only and since_iso", async () => {
    const { bridge, call } = setup();
    const now = Date.now();
    const read = bridge.seedMail({ subject: "read", is_read: true, received_at: new Date(now - 1_000).toISOString() });
    bridge.seedMail({ subject: "unread", is_read: false, received_at: new Date(now - 60_000).toISOString() });
    const unread = await call("list_mail", { unread_only: true });
    expect((unread.data as { items: { subject: string }[] }).items.map((m) => m.subject)).toEqual(["unread"]);
    const recent = await call("list_mail", { since_iso: new Date(now - 10_000).toISOString() });
    expect((recent.data as { items: { subject: string }[] }).items.map((m) => m.subject)).toEqual(["read"]);
    expect(read.entry_id).toBeTruthy();
  });

  test("list_mail rejects an invalid since_iso", async () => {
    const { call } = setup();
    const result = await call("list_mail", { since_iso: "yesterday" });
    expect(result.isError).toBe(true);
    expect((result.data as { error: string }).error).toBe("INVALID_ARGUMENT");
  });

  test("search_mail supports subject, sender and subject_body", async () => {
    const { bridge, call } = setup();
    bridge.seedMail({ subject: "Budget review", body: "numbers", sender_smtp: "alice@example.com" });
    bridge.seedMail({ subject: "Lunch", body: "budget talk", sender_smtp: "bob@example.com" });
    const subject = await call("search_mail", { query: "Budget", scope: "subject" });
    expect((subject.data as { count: number }).count).toBe(1);
    const sender = await call("search_mail", { query: "alice", scope: "sender" });
    expect((sender.data as { count: number }).count).toBe(1);
    const both = await call("search_mail", { query: "budget", scope: "subject_body" });
    expect((both.data as { count: number }).count).toBe(2);
  });

  test("read_mail truncates by UTF-8 bytes without splitting a character", async () => {
    const { bridge, call } = setup();
    const mail = bridge.seedMail({ subject: "unicode", body: "ääääääääää" });
    const result = await call("read_mail", { entry_id: mail.entry_id, max_body_bytes: 3 });
    const data = result.data as { body: string; body_truncated: boolean };
    expect(data.body_truncated).toBe(true);
    expect(Buffer.byteLength(data.body, "utf8")).toBeLessThanOrEqual(3);
    expect(data.body).toBe("ä");
  });

  test("read_mail includes recipients and attachment metadata", async () => {
    const { bridge, call } = setup();
    const mail = bridge.seedMail({
      subject: "with attachment",
      to: "one@example.com; two@example.com",
      cc: "cc@example.com",
      attachments: [{ filename: "report.txt", size: 12, content: new TextEncoder().encode("hello") }],
    });
    const result = await call("read_mail", { entry_id: mail.entry_id });
    const data = result.data as {
      recipients: { smtp: string; type: number }[];
      attachments: { index: number; filename: string }[];
    };
    expect(data.recipients.map((r) => r.smtp)).toEqual(["one@example.com", "two@example.com", "cc@example.com"]);
    expect(data.attachments[0]).toMatchObject({ index: 1, filename: "report.txt" });
  });

  test("download_attachment writes the file", async () => {
    const { bridge, dir, call } = setup();
    const mail = bridge.seedMail({
      attachments: [{ filename: "note.txt", size: 5, content: new TextEncoder().encode("hello") }],
    });
    const dest = join(dir, "downloads");
    const result = await call("download_attachment", { entry_id: mail.entry_id, index: 1, dest_dir: dest });
    const saved = (result.data as { saved_to: string }).saved_to;
    expect(existsSync(saved)).toBe(true);
    expect(readFileSync(saved, "utf8")).toBe("hello");
  });

  test("whoami lists the default store and accounts", async () => {
    const { call } = setup();
    const result = await call("whoami");
    const data = result.data as {
      default_store: string;
      accounts: { smtp: string }[];
      accessible_mailboxes: { is_default: boolean }[];
    };
    expect(data.default_store).toBe("Personal Mailbox");
    expect(data.accounts[0]!.smtp).toBe("user@example.com");
    expect(data.accessible_mailboxes[0]!.is_default).toBe(true);
  });

  test("guardrails_status shows sending off by default", async () => {
    const { call } = setup();
    const data = (await call("guardrails_status")).data as {
      allow_send_direct: boolean;
      unrestricted: boolean;
      allowed_recipients: string[];
    };
    expect(data.allow_send_direct).toBe(false);
    expect(data.unrestricted).toBe(true);
    expect(data.allowed_recipients).toEqual([]);
  });

  test("health_check reports OK and then ERROR without throwing", async () => {
    const { bridge, call } = setup();
    const good = (await call("health_check")).data as { status: string; outlook_running: boolean };
    expect(good.status).toBe("OK");
    expect(good.outlook_running).toBe(true);
    bridge.closeOutlook();
    const bad = (await call("health_check")).data as { status: string; error: string };
    expect(bad.status).toBe("ERROR");
    expect(bad.error).toContain("Outlook");
  });

  test("list_events and read_event return the calendar window", async () => {
    const { bridge, call } = setup();
    const event = bridge.seedEvent({
      subject: "Planning",
      start: new Date(Date.now() + 3_600_000).toISOString(),
      required: "one@example.com",
    });
    const list = await call("list_events", { days_ahead: 2 });
    const data = list.data as { count: number; events: { subject: string }[] };
    expect(data.count).toBe(1);
    expect(data.events[0]!.subject).toBe("Planning");
    const read = await call("read_event", { entry_id: event.entry_id });
    const detail = read.data as { subject: string; recipients: { smtp: string }[] };
    expect(detail.subject).toBe("Planning");
    expect(detail.recipients[0]!.smtp).toBe("one@example.com");
  });

  test("find_freeslots skips busy time and respects business hours", async () => {
    const monday = new Date(2026, 3, 20, 9, 0, 0);
    const { bridge, call } = setup({}, monday);
    bridge.seedEvent({
      subject: "busy",
      start: new Date(2026, 3, 20, 10, 0, 0).toISOString(),
      end: new Date(2026, 3, 20, 11, 0, 0).toISOString(),
    });
    const result = await call("find_freeslots", {
      duration_minutes: 30,
      days_ahead: 1,
      business_hours: "09:00-18:00",
      weekdays_only: true,
      slot_step_minutes: 30,
      limit: 4,
    });
    const data = result.data as { slots: { start: string; end: string }[]; busy_count: number };
    expect(data.busy_count).toBe(1);
    expect(data.slots.map((s) => s.start)).toEqual([
      "2026-04-20T09:00",
      "2026-04-20T09:30",
      "2026-04-20T11:00",
      "2026-04-20T11:30",
    ]);
  });

  test("find_freeslots_multi merges attendee free/busy", async () => {
    const monday = new Date(2026, 3, 20, 9, 0, 0);
    const { bridge, call } = setup({}, monday);
    const busy = Array.from({ length: 48 }, (_, i) => (i === 20 || i === 21 ? "2" : "0")).join("");
    bridge.state.attendee_busy.set("bob@example.com", busy);
    const result = await call("find_freeslots_multi", {
      attendees: "bob@example.com",
      duration_minutes: 30,
      days_ahead: 1,
      business_hours: "09:00-18:00",
      slot_step_minutes: 30,
      limit: 4,
      include_self: false,
    });
    const data = result.data as {
      attendee_status: Record<string, { busy_intervals: number; error: string | null }>;
      slots: { start: string }[];
    };
    expect(data.attendee_status["bob@example.com"]!.busy_intervals).toBe(1);
    expect(data.slots.map((s) => s.start)).toEqual([
      "2026-04-20T09:00",
      "2026-04-20T09:30",
      "2026-04-20T11:00",
      "2026-04-20T11:30",
    ]);
  });

  test("an unknown tool is an argument error", async () => {
    const { call } = setup();
    const result = await call("nope");
    expect(result.isError).toBe(true);
    expect((result.data as { error: string }).error).toBe("INVALID_ARGUMENT");
  });
});
