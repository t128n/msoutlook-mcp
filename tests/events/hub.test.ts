import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { defineEvent } from "../../packages/events/src/index.ts";
import { FakeComBridge } from "../../src/com/fake.ts";
import { EventHub, EventHubError } from "../../src/events/hub.ts";
import { createNewMailEvent } from "../../src/events/new-mail.ts";
import { EventStateStore } from "../../src/events/state.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

function tempState() {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-events-"));
  dirs.push(dir);
  return join(dir, "events.json");
}

function tickEvent(counter: { n: number }, options: { failFirst?: boolean } = {}) {
  return defineEvent<{ value?: string }, { n: number }>({
    name: "tick",
    description: "test tick",
    delivery: ["poll", "push"],
    inputSchema: { type: "object", properties: {} },
    payloadSchema: { type: "object", properties: { n: { type: "integer" } } },
    async poll(cursor) {
      counter.n += 1;
      if (options.failFirst && counter.n === 1) throw new Error("outlook is away");
      if (cursor === null) {
        return {
          events: [{ eventId: "tick-0", name: "tick", timestamp: new Date().toISOString(), data: { n: 0 } }],
          cursor: "0",
          hasMore: false,
          nextPollMs: 1_000,
        };
      }
      const n = Number(cursor) + 1;
      return {
        events: [{ eventId: `tick-${n}`, name: "tick", timestamp: new Date().toISOString(), data: { n } }],
        cursor: String(n),
        hasMore: false,
        nextPollMs: 1_000,
      };
    },
  });
}

async function collect(stream: ReadableStream, ms: number): Promise<unknown[]> {
  const reader = stream.getReader();
  const out: unknown[] = [];
  let timer: ReturnType<typeof setTimeout> | undefined;
  const deadline = new Promise<void>((resolve) => {
    timer = setTimeout(() => resolve(), ms);
  });
  const reading = (async () => {
    try {
      for (;;) {
        const next = await reader.read();
        if (next.done) break;
        out.push(next.value);
      }
    } catch {
      /* cancelled */
    }
  })();
  await Promise.race([reading, deadline]);
  if (timer) clearTimeout(timer);
  try {
    await reader.cancel();
  } catch {
    /* already closed */
  }
  return out;
}

describe("event hub", () => {
  test("the persisted cursor survives a restart without replay or skip", async () => {
    const file = tempState();
    let now = new Date("2026-04-20T09:00:00.000Z");
    const bridge = new FakeComBridge();
    const make = () =>
      new EventHub({
        definitions: [createNewMailEvent({ bridge, intervalMs: 1_000, now: () => now })],
        store: new EventStateStore(file, () => now),
        now: () => now,
      });

    const first = make();
    const initial = await first.poll("new_mail", {}, null);
    first.close();

    now = new Date(now.getTime() + 5_000);
    const oldMail = bridge.seedMail({ subject: "old", received_at: now.toISOString() });
    const after = make();
    const delivered = await after.poll("new_mail", {}, null);
    expect(delivered.events.map((e) => (e.data as { entryId: string }).entryId)).toEqual([oldMail.entry_id]);
    after.close();

    now = new Date(now.getTime() + 5_000);
    const newMail = bridge.seedMail({ subject: "new", received_at: now.toISOString() });
    const restarted = make();
    const replayed = await restarted.poll("new_mail", {}, null);
    expect(replayed.events.map((e) => (e.data as { entryId: string }).entryId)).toEqual([newMail.entry_id]);
    restarted.close();
  });

  test("two streams for the same subscription share one poll loop", async () => {
    const counter = { n: 0 };
    const hub = new EventHub({
      definitions: [tickEvent(counter)],
      store: new EventStateStore(tempState()),
    });
    const a = await hub.openStream("tick", {}, null, 1_000);
    const b = await hub.openStream("tick", {}, null, 1_000);
    const [am, bm] = await Promise.all([collect(a.stream.messages, 80), collect(b.stream.messages, 80)]);
    a.stream.cancel();
    b.stream.cancel();
    hub.close();
    expect(counter.n).toBeLessThanOrEqual(3);
    expect(counter.n).toBeGreaterThanOrEqual(2);
    expect(am.some((m) => (m as { type: string }).type === "active")).toBe(true);
    expect(bm.some((m) => (m as { type: string }).type === "active")).toBe(true);
    expect(am.some((m) => (m as { type: string }).type === "event")).toBe(true);
    expect(bm.some((m) => (m as { type: string }).type === "event")).toBe(true);
  });

  test("a failing poll sends an error and keeps the stream alive", async () => {
    const counter = { n: 0 };
    const hub = new EventHub({
      definitions: [tickEvent(counter, { failFirst: true })],
      store: new EventStateStore(tempState()),
      backoffInitialMs: 10,
      backoffMaxMs: 20,
    });
    const opened = await hub.openStream("tick", {}, null, 1_000);
    const messages = await collect(opened.stream.messages, 150);
    opened.stream.cancel();
    hub.close();
    const types = messages.map((m) => (m as { type: string }).type);
    expect(types).toContain("active");
    expect(types).toContain("error");
    expect(types).toContain("event");
    const error = messages.find((m) => (m as { type: string }).type === "error") as {
      error: { message: string; data: { reason: string } };
    };
    expect(error.error.message).toBe("UpstreamError");
    expect(error.error.data.reason).toContain("outlook is away");
  });

  test("a heartbeat is sent during a quiet stream", async () => {
    const counter = { n: 0 };
    const hub = new EventHub({
      definitions: [
        defineEvent({
          name: "quiet",
          description: "quiet",
          delivery: ["poll", "push"],
          inputSchema: { type: "object", properties: {} },
          payloadSchema: { type: "object", properties: {} },
          async poll() {
            counter.n += 1;
            return { events: [], cursor: "0", hasMore: false, nextPollMs: 60_000 };
          },
        }),
      ],
      store: new EventStateStore(tempState()),
      heartbeatMs: 1_000,
    });
    const opened = await hub.openStream("quiet", {}, null, 1_000);
    const messages = await collect(opened.stream.messages, 1_300);
    opened.stream.cancel();
    hub.close();
    expect(messages.some((m) => (m as { type: string }).type === "heartbeat")).toBe(true);
  });

  test("the subscription cap is enforced", async () => {
    const counter = { n: 0 };
    const hub = new EventHub({
      definitions: [tickEvent(counter)],
      store: new EventStateStore(tempState()),
      maxSubscriptions: 1,
    });
    await hub.openStream("tick", { value: "a" }, null, 1_000);
    await expect(hub.openStream("tick", { value: "b" }, null, 1_000)).rejects.toBeInstanceOf(EventHubError);
    hub.close();
  });

  test("an unknown event is not_found", async () => {
    const hub = new EventHub({ definitions: [], store: new EventStateStore(tempState()) });
    await expect(hub.poll("nope", {}, null)).rejects.toMatchObject({ kind: "not_found" });
    await expect(hub.openStream("nope", {}, null)).rejects.toMatchObject({ kind: "not_found" });
    hub.close();
  });

  test("invalid arguments are invalid_params", async () => {
    const hub = new EventHub({
      definitions: [createNewMailEvent({ bridge: new FakeComBridge() })],
      store: new EventStateStore(tempState()),
    });
    await expect(
      hub.poll("new_mail", { unreadOnly: "not a boolean" } as unknown as Record<string, unknown>, null),
    ).rejects.toMatchObject({ kind: "invalid_params" });
    hub.close();
  });

  test("a payload over the cap is dropped and marked truncated", async () => {
    const huge = "x".repeat(70_000);
    const hub = new EventHub({
      definitions: [
        defineEvent({
          name: "huge",
          description: "huge",
          delivery: ["poll"],
          inputSchema: { type: "object", properties: {} },
          payloadSchema: { type: "object", properties: {} },
          async poll() {
            return {
              events: [{ eventId: "huge-1", name: "huge", timestamp: new Date().toISOString(), data: { huge } }],
              cursor: "1",
              hasMore: false,
              nextPollMs: 1_000,
            };
          },
        }),
      ],
      store: new EventStateStore(tempState()),
    });
    const result = await hub.poll("huge", {}, null);
    expect(result.events).toEqual([]);
    expect(result.truncated).toBe(true);
    hub.close();
  });
});
