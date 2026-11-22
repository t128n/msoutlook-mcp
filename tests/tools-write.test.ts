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
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-write-"));
  dirs.push(dir);
  const paths = appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin");
  const bridge = new FakeComBridge();
  const guardrails = new Guardrails({ ...DEFAULT_CONFIG, ...config });
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
  return { dir, paths, bridge, guardrails, tools, call };
}

const allowed = { allowed_recipients: ["example.com"] };

describe("guarded mail writes", () => {
  test("create_draft saves a draft and never sends", async () => {
    const { bridge, call } = setup(allowed);
    const result = await call("create_draft", {
      to: "one@example.com",
      cc: "two@example.com",
      bcc: "three@example.com",
      subject: "Hello",
      body: "Body",
    });
    expect(result.isError).toBeUndefined();
    const data = result.data as { entry_id: string; to: string[]; cc: string[]; bcc: string[] };
    expect(data.to).toEqual(["one@example.com"]);
    expect(data.cc).toEqual(["two@example.com"]);
    expect(data.bcc).toEqual(["three@example.com"]);
    const mail = bridge.state.mail.get(data.entry_id)!;
    expect(mail.folder).toBe("drafts");
    expect(mail.subject).toBe("Hello");
  });

  test("create_draft refuses an outside recipient before writing", async () => {
    const { bridge, call } = setup(allowed);
    const result = await call("create_draft", {
      to: "outside@example.org",
      subject: "Hello",
      body: "Body",
    });
    expect(result.isError).toBe(true);
    expect((result.data as { error: string }).error).toBe("RECIPIENTS_NOT_ALLOWED");
    expect((result.data as { rejected: string[] }).rejected).toEqual(["outside@example.org"]);
    expect(bridge.state.mail.size).toBe(0);
  });

  test("an empty allowlist means no restriction", async () => {
    const { call } = setup();
    const result = await call("create_draft", {
      to: "anyone@anywhere.example",
      subject: "Hello",
      body: "Body",
    });
    expect(result.isError).toBeUndefined();
  });

  test("from_mailbox must be in shared_mailboxes", async () => {
    const { call } = setup({ ...allowed, shared_mailboxes: [] });
    const result = await call("create_draft", {
      to: "one@example.com",
      subject: "Hello",
      body: "Body",
      from_mailbox: "team@example.com",
    });
    expect(result.isError).toBe(true);
    expect((result.data as { error: string }).error).toBe("FROM_MAILBOX_NOT_ALLOWED");
  });

  test("an allowed from_mailbox moves the draft to its Drafts folder", async () => {
    const { bridge, call } = setup({ ...allowed, shared_mailboxes: ["team@example.com"] });
    const result = await call("create_draft", {
      to: "one@example.com",
      subject: "Hello",
      body: "Body",
      from_mailbox: "team@example.com",
    });
    const data = result.data as { entry_id: string; drafts_folder: string };
    expect(data.drafts_folder).toContain("Drafts");
    expect(bridge.state.mail.get(data.entry_id)!.folder).toContain("team@example.com");
  });

  test("send_draft needs confirm, the send switch and the allowlist", async () => {
    const { bridge, call } = setup(allowed);
    const draft = (await call("create_draft", { to: "one@example.com", subject: "x", body: "y" })).data as {
      entry_id: string;
    };
    const noConfirm = await call("send_draft", { entry_id: draft.entry_id });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    const noSend = await call("send_draft", { entry_id: draft.entry_id, confirm: true });
    expect((noSend.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    expect(bridge.state.mail.get(draft.entry_id)!.folder).toBe("drafts");
  });

  test("a draft edited to add an outside recipient is refused at send time", async () => {
    const { bridge, call } = setup({ ...allowed, allow_send: true });
    const draft = (await call("create_draft", { to: "one@example.com", subject: "x", body: "y" })).data as {
      entry_id: string;
    };
    bridge.state.mail.get(draft.entry_id)!.to = "outside@example.org";
    const result = await call("send_draft", { entry_id: draft.entry_id, confirm: true });
    expect((result.data as { error: string }).error).toBe("RECIPIENTS_NOT_ALLOWED_AT_SEND");
    expect(bridge.state.mail.get(draft.entry_id)!.folder).toBe("drafts");
  });

  test("send_draft sends when every guardrail passes", async () => {
    const { bridge, call } = setup({ ...allowed, allow_send: true });
    const draft = (await call("create_draft", { to: "one@example.com", subject: "x", body: "y" })).data as {
      entry_id: string;
    };
    const result = await call("send_draft", { entry_id: draft.entry_id, confirm: true });
    expect(result.isError).toBeUndefined();
    expect(bridge.state.mail.get(draft.entry_id)!.folder).toBe("sent");
  });

  test("reply_mail saves a draft by default and reports outside recipients", async () => {
    const { bridge, call } = setup(allowed);
    const original = bridge.seedMail({ sender_smtp: "outside@example.org", subject: "Question" });
    const result = await call("reply_mail", { entry_id: original.entry_id, body: "Answer" });
    const data = result.data as {
      draft_entry_id: string;
      recipients: string[];
      recipients_outside_allowlist: string[];
    };
    expect(data.recipients).toEqual(["outside@example.org"]);
    expect(data.recipients_outside_allowlist).toEqual(["outside@example.org"]);
    expect(bridge.state.mail.get(data.draft_entry_id)!.folder).toBe("drafts");
  });

  test("reply_mail direct send needs confirm and sending enabled", async () => {
    const { bridge, call } = setup(allowed);
    const original = bridge.seedMail({ sender_smtp: "one@example.com", subject: "Question" });
    const noConfirm = await call("reply_mail", {
      entry_id: original.entry_id,
      body: "Answer",
      save_only: false,
    });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");

    const noSend = await call("reply_mail", {
      entry_id: original.entry_id,
      body: "Answer",
      save_only: false,
      confirm: true,
    });
    expect((noSend.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    expect((noSend.data as { saved_draft_entry_id: string }).saved_draft_entry_id).toBeTruthy();

    const enabled = setup({ ...allowed, allow_send: true });
    const original2 = enabled.bridge.seedMail({ sender_smtp: "one@example.com", subject: "Question" });
    const sent = await enabled.call("reply_mail", {
      entry_id: original2.entry_id,
      body: "Answer",
      save_only: false,
      confirm: true,
    });
    expect(sent.isError).toBeUndefined();
    expect((sent.data as { sent: boolean }).sent).toBe(true);
  });

  test("reply_mail with from_mailbox always saves a draft", async () => {
    const { bridge, call } = setup({ ...allowed, allow_send: true, shared_mailboxes: ["team@example.com"] });
    const original = bridge.seedMail({ sender_smtp: "one@example.com", subject: "Question" });
    const result = await call("reply_mail", {
      entry_id: original.entry_id,
      body: "Answer",
      save_only: false,
      confirm: true,
      from_mailbox: "team@example.com",
    });
    expect((result.data as { error: string }).error).toBe("FROM_MAILBOX_REQUIRES_SAVE_ONLY");
    expect((result.data as { saved_draft_entry_id: string }).saved_draft_entry_id).toBeTruthy();
  });

  test("move_mail, mark_read and flag_mail need confirm", async () => {
    const { bridge, call } = setup();
    const mail = bridge.seedMail({});
    for (const [name, args] of [
      ["move_mail", { entry_id: mail.entry_id, target_folder: "archive" }],
      ["mark_read", { entry_id: mail.entry_id, read: true }],
      ["flag_mail", { entry_id: mail.entry_id, flag: true }],
    ] as [string, Record<string, unknown>][]) {
      const result = await call(name, args);
      expect((result.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    }
    expect((await call("move_mail", { entry_id: mail.entry_id, target_folder: "archive", confirm: true })).isError).toBeUndefined();
    expect((await call("mark_read", { entry_id: mail.entry_id, read: true, confirm: true })).isError).toBeUndefined();
    expect((await call("flag_mail", { entry_id: mail.entry_id, flag: true, confirm: true })).isError).toBeUndefined();
    expect(bridge.state.mail.size).toBe(1);
    expect([...bridge.state.mail.values()][0]!.flagged).toBe(true);
  });
});

describe("guarded calendar writes", () => {
  test("create_event_draft checks attendees and does not invite", async () => {
    const { bridge, call } = setup(allowed);
    const blocked = await call("create_event_draft", {
      subject: "Planning",
      start_iso: "2026-04-21T10:00:00",
      required_attendees: "outside@example.org",
    });
    expect((blocked.data as { error: string }).error).toBe("RECIPIENTS_NOT_ALLOWED");
    expect(bridge.state.events.size).toBe(0);

    const result = await call("create_event_draft", {
      subject: "Planning",
      start_iso: "2026-04-21T10:00:00",
      duration_minutes: 45,
      required_attendees: "one@example.com",
      optional_attendees: ["two@example.com"],
    });
    const data = result.data as { entry_id: string; is_meeting: boolean; required: string[] };
    expect(data.is_meeting).toBe(true);
    expect(data.required).toEqual(["one@example.com"]);
    expect(bridge.state.events.get(data.entry_id)!.invite_sent).toBe(false);
  });

  test("send_event_invites needs confirm, the switch and a fresh allowlist check", async () => {
    const { bridge, call } = setup({ ...allowed, allow_send: false });
    const event = (await call("create_event_draft", {
      subject: "Planning",
      start_iso: "2026-04-21T10:00:00",
      required_attendees: "one@example.com",
    })).data as { entry_id: string };

    const noConfirm = await call("send_event_invites", { entry_id: event.entry_id });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    const noSend = await call("send_event_invites", { entry_id: event.entry_id, confirm: true });
    expect((noSend.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");

    const enabled = setup({ ...allowed, allow_send: true });
    const event2 = (await enabled.call("create_event_draft", {
      subject: "Planning",
      start_iso: "2026-04-21T10:00:00",
      required_attendees: "one@example.com",
    })).data as { entry_id: string };
    enabled.bridge.state.events.get(event2.entry_id)!.required = "outside@example.org";
    const blocked = await enabled.call("send_event_invites", { entry_id: event2.entry_id, confirm: true });
    expect((blocked.data as { error: string }).error).toBe("RECIPIENTS_NOT_ALLOWED_AT_SEND");

    enabled.bridge.state.events.get(event2.entry_id)!.required = "one@example.com";
    const sent = await enabled.call("send_event_invites", { entry_id: event2.entry_id, confirm: true });
    expect(sent.isError).toBeUndefined();
    expect(enabled.bridge.state.events.get(event2.entry_id)!.invite_sent).toBe(true);
    expect(bridge.state.events.get(event.entry_id)!.invite_sent).toBe(false);
  });

  test("respond_meeting needs confirm; notifying the organizer needs sending", async () => {
    const { bridge, call } = setup();
    const event = bridge.seedEvent({
      subject: "Planning",
      start: new Date(Date.now() + 3_600_000).toISOString(),
      meeting_status: 1,
    });
    const noConfirm = await call("respond_meeting", { entry_id: event.entry_id, response: "accepted" });
    expect((noConfirm.data as { error: string }).error).toBe("CONFIRM_REQUIRED");
    const noSend = await call("respond_meeting", {
      entry_id: event.entry_id,
      response: "accepted",
      send_response: true,
      confirm: true,
    });
    expect((noSend.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    const saved = await call("respond_meeting", {
      entry_id: event.entry_id,
      response: "tentative",
      confirm: true,
    });
    expect(saved.isError).toBeUndefined();
    expect((saved.data as { action: string }).action).toBe("saved");
    expect(bridge.state.events.get(event.entry_id)!.response_status).toBe(2);
  });

  test("turning sending on takes effect without rebuilding the tool", async () => {
    let config: AppConfig = { ...DEFAULT_CONFIG, allowed_recipients: ["example.com"], allow_send: false };
    const dir = mkdtempSync(join(tmpdir(), "msoutlook-write-toggle-"));
    dirs.push(dir);
    const bridge = new FakeComBridge();
    const tools = createTools({
      bridge,
      guardrails: new Guardrails(() => config),
      health: new HealthMonitor({ bridge, version: "0.1.0", cacheMs: 0, timeoutMs: 200 }),
      paths: appPaths({ MSOUTLOOK_MCP_HOME: dir }, "darwin"),
      version: "0.1.0",
      logger: new Logger(join(dir, "server.log"), "error"),
      now: () => new Date(),
    });
    const call = async (name: string, args: Record<string, unknown> = {}) =>
      (await tools.call(name, args)) as ToolResponse;
    const draft = (await call("create_draft", { to: "one@example.com", subject: "x", body: "y" })).data as {
      entry_id: string;
    };
    const off = await call("send_draft", { entry_id: draft.entry_id, confirm: true });
    expect((off.data as { error: string }).error).toBe("ALLOW_SEND_DISABLED");
    config = { ...config, allow_send: true };
    const on = await call("send_draft", { entry_id: draft.entry_id, confirm: true });
    expect(on.isError).toBeUndefined();
    expect(bridge.state.mail.get(draft.entry_id)!.folder).toBe("sent");
  });

  test("the tool registry now has 21 tools", async () => {
    const { tools } = setup();
    expect(tools.list()).toHaveLength(21);
    expect(tools.list().map((t) => t.name)).toContain("send_draft");
  });
});
