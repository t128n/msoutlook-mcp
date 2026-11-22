import { describe, expect, test } from "bun:test";
import { FakeComBridge } from "../../src/com/fake.ts";
import { createNewMailEvent, decodeCursor, encodeCursor } from "../../src/events/new-mail.ts";

function setup(start = new Date("2026-04-20T09:00:00.000Z")) {
  let now = start;
  const bridge = new FakeComBridge();
  const event = createNewMailEvent({ bridge, intervalMs: 30_000, now: () => now });
  const clock = {
    get: () => now,
    set: (value: Date) => {
      now = value;
    },
    advance: (ms: number) => {
      now = new Date(now.getTime() + ms);
    },
  };
  return { bridge, event, clock };
}

async function firstCursor(event: ReturnType<typeof createNewMailEvent>) {
  const first = await event.poll(null, {}, { limit: 50 });
  expect(first.events).toEqual([]);
  return first.cursor;
}

describe("new_mail event", () => {
  test("the first poll starts from now and emits nothing", async () => {
    const { event } = setup();
    const first = await event.poll(null, {}, { limit: 50 });
    expect(first.events).toEqual([]);
    expect(first.hasMore).toBe(false);
    expect(first.nextPollMs).toBe(30_000);
    expect(decodeCursor(first.cursor)?.at).toBe("2026-04-20T09:00:00.000Z");
  });

  test("a new mail produces one thin event", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(10_000);
    const mail = bridge.seedMail({
      subject: "Quarterly report",
      body: "the body must never be in the event",
      received_at: clock.get().toISOString(),
      sender_smtp: "alice@example.com",
      internet_message_id: "<message-1@example.com>",
      attachments: [{ filename: "report.pdf", size: 10 }],
      is_read: false,
    });
    const result = await event.poll(cursor, {}, { limit: 50 });
    expect(result.events).toHaveLength(1);
    const occurrence = result.events[0]!;
    expect(occurrence.eventId).toBe("<message-1@example.com>");
    expect(occurrence.name).toBe("new_mail");
    expect(occurrence.timestamp).toBe(clock.get().toISOString());
    expect(occurrence.data).toEqual({
      entryId: mail.entry_id,
      folder: "inbox",
      senderSmtp: "alice@example.com",
      subject: "Quarterly report",
      receivedAt: clock.get().toISOString(),
      hasAttachments: true,
      unread: true,
    });
    expect(JSON.stringify(occurrence.data)).not.toContain("the body");
  });

  test("repeating a poll with the new cursor does not duplicate", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    bridge.seedMail({ received_at: clock.get().toISOString() });
    const first = await event.poll(cursor, {}, { limit: 50 });
    expect(first.events).toHaveLength(1);
    const again = await event.poll(first.cursor, {}, { limit: 50 });
    expect(again.events).toEqual([]);
  });

  test("mails with the same ReceivedTime use the id tie-breaker", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    bridge.seedMail({ subject: "one", received_at: clock.get().toISOString(), internet_message_id: "<one@example.com>" });
    bridge.seedMail({ subject: "two", received_at: clock.get().toISOString(), internet_message_id: "<two@example.com>" });
    const first = await event.poll(cursor, {}, { limit: 50 });
    expect(first.events.map((e) => e.eventId).sort()).toEqual(["<one@example.com>", "<two@example.com>"]);
    const again = await event.poll(first.cursor, {}, { limit: 50 });
    expect(again.events).toEqual([]);

    clock.advance(1_000);
    bridge.seedMail({ subject: "three", received_at: clock.get().toISOString(), internet_message_id: "<three@example.com>" });
    const third = await event.poll(again.cursor, {}, { limit: 50 });
    expect(third.events.map((e) => e.data.subject)).toEqual(["three"]);
  });

  test("moving a delivered mail does not produce a second event", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    const mail = bridge.seedMail({ received_at: clock.get().toISOString() });
    const first = await event.poll(cursor, {}, { limit: 50 });
    expect(first.events).toHaveLength(1);
    await bridge.moveMail({ entryId: mail.entry_id, targetFolder: "archive" });
    const again = await event.poll(first.cursor, {}, { limit: 50 });
    expect(again.events).toEqual([]);
  });

  test("fromDomains filters senders", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    bridge.seedMail({ sender_smtp: "inside@example.com", received_at: clock.get().toISOString() });
    bridge.seedMail({ sender_smtp: "outside@example.org", received_at: clock.get().toISOString() });
    const result = await event.poll(cursor, { fromDomains: ["example.com"] }, { limit: 50 });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]!.data.senderSmtp).toBe("inside@example.com");
  });

  test("unreadOnly skips read mail but advances past it", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    bridge.seedMail({ subject: "read", is_read: true, received_at: clock.get().toISOString() });
    const skipped = await event.poll(cursor, { unreadOnly: true }, { limit: 50 });
    expect(skipped.events).toEqual([]);
    clock.advance(1_000);
    bridge.seedMail({ subject: "unread", is_read: false, received_at: clock.get().toISOString() });
    const result = await event.poll(skipped.cursor, { unreadOnly: true }, { limit: 50 });
    expect(result.events.map((e) => e.data.subject)).toEqual(["unread"]);
  });

  test("an invalid cursor resets to now and says truncated", async () => {
    const { event } = setup();
    const result = await event.poll("not-a-cursor", {}, { limit: 50 });
    expect(result.events).toEqual([]);
    expect(result.truncated).toBe(true);
    expect(decodeCursor(result.cursor)).not.toBeNull();
  });

  test("the batch cap returns hasMore and the next poll drains it", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(1_000);
    for (let i = 0; i < 55; i++) {
      bridge.seedMail({
        subject: `mail ${i}`,
        received_at: new Date(clock.get().getTime() + i * 1000).toISOString(),
        internet_message_id: `<m${i}@example.com>`,
      });
    }
    const first = await event.poll(cursor, {}, { limit: 50 });
    expect(first.events).toHaveLength(50);
    expect(first.hasMore).toBe(true);
    expect(first.truncated).toBe(false);
    const second = await event.poll(first.cursor, {}, { limit: 50 });
    expect(second.events).toHaveLength(5);
    expect(second.hasMore).toBe(false);
  });

  test("mail that arrives while Outlook is away is delivered on resume, once", async () => {
    const { bridge, event, clock } = setup();
    const cursor = await firstCursor(event);
    clock.advance(5_000);
    const mail = bridge.seedMail({ received_at: clock.get().toISOString() });
    bridge.closeOutlook();
    await expect(event.poll(cursor, {}, { limit: 50 })).rejects.toThrow();
    bridge.openOutlook();
    const resumed = await event.poll(cursor, {}, { limit: 50 });
    expect(resumed.events.map((e) => e.data.entryId)).toEqual([mail.entry_id]);
    const again = await event.poll(resumed.cursor, {}, { limit: 50 });
    expect(again.events).toEqual([]);
  });

  test("the cursor round-trips and survives a restart", async () => {
    const state = { v: 1 as const, at: "2026-04-20T09:00:00.000Z", ids: ["a", "b"] };
    expect(decodeCursor(encodeCursor(state))).toEqual(state);
    expect(decodeCursor("garbage")).toBeNull();
  });
});
