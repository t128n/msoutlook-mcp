import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../../src/com/fake.ts";
import { EventHub } from "../../src/events/hub.ts";
import { createNewMailEvent } from "../../src/events/new-mail.ts";
import { EventStateStore } from "../../src/events/state.ts";
import { HealthMonitor } from "../../src/server/health.ts";
import { startMcpServer, type McpServerHandle } from "../../src/server/mcp.ts";
import { Logger } from "../../src/util/logger.ts";

const handles: McpServerHandle[] = [];
const dirs: string[] = [];
afterEach(() => {
  for (const handle of handles.splice(0)) handle.stop();
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

function start() {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-events-http-"));
  dirs.push(dir);
  const bridge = new FakeComBridge();
  const health = new HealthMonitor({ bridge, version: "0.1.0", cacheMs: 0, timeoutMs: 500 });
  const hub = new EventHub({
    definitions: [createNewMailEvent({ bridge, intervalMs: 20 })],
    store: new EventStateStore(join(dir, "events.json")),
    minNextPollMs: 10,
    onPoll: () => health.markPoll(),
  });
  const server = startMcpServer({
    config: { port: 0 },
    token: "test-token",
    health,
    version: "0.1.0",
    logger: new Logger(join(dir, "server.log"), "error"),
    events: hub,
  });
  handles.push(server);
  const base = `http://127.0.0.1:${server.port}`;
  const auth = { authorization: "Bearer test-token", "content-type": "application/json" };
  const eventHeaders = { ...auth, "mcp-name": "new_mail" };
  const rpc = async (body: unknown, headers: Record<string, string> = auth) =>
    fetch(`${base}/mcp`, { method: "POST", headers, body: JSON.stringify(body) });
  return { server, base, auth, eventHeaders, rpc, bridge, hub, health };
}

interface Frame {
  jsonrpc?: string;
  id?: unknown;
  method?: string;
  params?: Record<string, unknown>;
  result?: Record<string, unknown>;
  error?: { code: number; message: string };
}

async function readFrames(response: Response, timeoutMs: number, stopAfter?: (frame: Frame) => boolean): Promise<Frame[]> {
  const reader = response.body!.getReader();
  const decoder = new TextDecoder();
  const frames: Frame[] = [];
  let buffer = "";
  const deadline = Date.now() + timeoutMs;
  try {
    for (;;) {
      if (Date.now() > deadline) break;
      const next = await Promise.race([
        reader.read(),
        new Promise<{ done: true; value: undefined }>((resolve) =>
          setTimeout(() => resolve({ done: true, value: undefined }), Math.max(0, deadline - Date.now())),
        ),
      ]);
      if (next.done) break;
      buffer += decoder.decode(next.value, { stream: true });
      let cut = buffer.indexOf("\n\n");
      while (cut >= 0) {
        const raw = buffer.slice(0, cut);
        buffer = buffer.slice(cut + 2);
        const data = raw
          .split("\n")
          .filter((line) => line.startsWith("data: "))
          .map((line) => line.slice(6))
          .join("");
        if (data) {
          const frame = JSON.parse(data) as Frame;
          frames.push(frame);
          if (stopAfter?.(frame)) return frames;
        }
        cut = buffer.indexOf("\n\n");
      }
    }
  } finally {
    try {
      await reader.cancel();
    } catch {
      /* already gone */
    }
  }
  return frames;
}

describe("MCP events wire format", () => {
  test("server/discover advertises the events extension", async () => {
    const { rpc } = start();
    const body = (await (await rpc({ jsonrpc: "2.0", id: 1, method: "server/discover" })).json()) as {
      result: { capabilities: { extensions: Record<string, unknown> } };
    };
    expect(body.result.capabilities.extensions["io.modelcontextprotocol/events"]).toBeDefined();
  });

  test("events/list returns new_mail with poll and push", async () => {
    const { rpc } = start();
    const body = (await (await rpc({ jsonrpc: "2.0", id: 2, method: "events/list" })).json()) as {
      result: { events: { name: string; delivery: string[]; inputSchema: unknown; payloadSchema: unknown }[] };
    };
    const event = body.result.events[0]!;
    expect(event.name).toBe("new_mail");
    expect(event.delivery).toEqual(["poll", "push"]);
    expect(event.inputSchema).toBeDefined();
    expect(event.payloadSchema).toBeDefined();
  });

  test("events/poll requires the Mcp-Name header", async () => {
    const { rpc } = start();
    const response = await rpc({
      jsonrpc: "2.0",
      id: 3,
      method: "events/poll",
      params: { name: "new_mail", cursor: null },
    });
    expect(response.status).toBe(400);
  });

  test("poll starts from now, then returns a thin event", async () => {
    const { rpc, eventHeaders, bridge } = start();
    const first = (await (
      await rpc({ jsonrpc: "2.0", id: 4, method: "events/poll", params: { name: "new_mail", arguments: {}, cursor: null } }, eventHeaders)
    ).json()) as { result: { events: unknown[]; cursor: string; hasMore: boolean; nextPollMs: number } };
    expect(first.result.events).toEqual([]);
    expect(first.result.hasMore).toBe(false);
    expect(typeof first.result.cursor).toBe("string");

    const mail = bridge.seedMail({
      subject: "Hello",
      body: "do not leak this body",
      received_at: new Date(Date.now() + 10).toISOString(),
      sender_smtp: "sender@example.com",
      internet_message_id: "<one@example.com>",
    });
    const second = (await (
      await rpc(
        {
          jsonrpc: "2.0",
          id: 5,
          method: "events/poll",
          params: { name: "new_mail", arguments: {}, cursor: first.result.cursor },
        },
        eventHeaders,
      )
    ).json()) as {
      result: {
        events: {
          eventId: string;
          name: string;
          timestamp: string;
          data: Record<string, unknown>;
        }[];
        cursor: string;
      };
    };
    expect(second.result.events).toHaveLength(1);
    const event = second.result.events[0]!;
    expect(event.eventId).toBe("<one@example.com>");
    expect(event.name).toBe("new_mail");
    expect(event.data.entryId).toBe(mail.entry_id);
    expect(event.data.subject).toBe("Hello");
    expect(JSON.stringify(event.data)).not.toContain("do not leak");
  });

  test("an unknown event is -32023 NotFound", async () => {
    const { rpc } = start();
    const body = (await (
      await rpc(
        { jsonrpc: "2.0", id: 6, method: "events/poll", params: { name: "nope", cursor: null } },
        { authorization: "Bearer test-token", "content-type": "application/json", "mcp-name": "nope" },
      )
    ).json()) as { error: { code: number } };
    expect(body.error.code).toBe(-32023);
  });

  test("events/stream sends active, then the event, with the subscription id", async () => {
    const { base, eventHeaders, rpc, bridge } = start();
    const first = (await (
      await rpc({ jsonrpc: "2.0", id: 7, method: "events/poll", params: { name: "new_mail", arguments: {}, cursor: null } }, eventHeaders)
    ).json()) as { result: { cursor: string } };
    bridge.seedMail({
      subject: "streamed",
      received_at: new Date(Date.now() + 10).toISOString(),
      internet_message_id: "<stream@example.com>",
    });
    const controller = new AbortController();
    const response = await fetch(`${base}/mcp`, {
      method: "POST",
      headers: eventHeaders,
      signal: controller.signal,
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: 7,
        method: "events/stream",
        params: { name: "new_mail", arguments: {}, cursor: first.result.cursor, heartbeatMs: 1000 },
      }),
    });
    expect(response.headers.get("content-type")).toContain("text/event-stream");
    const frames = await readFrames(response, 2_000, (frame) => frame.method === "notifications/events/event");
    controller.abort();
    expect(frames[0]!.method).toBe("notifications/events/active");
    expect(frames[0]!.params!.heartbeatMs).toBe(1000);
    const eventFrame = frames.find((frame) => frame.method === "notifications/events/event")!;
    expect(eventFrame.params!.eventId).toBe("<stream@example.com>");
    expect(eventFrame.params!._meta).toEqual({ "io.modelcontextprotocol/subscriptionId": 7 });
  });

  test("events/stream needs the Mcp-Name header and a known event", async () => {
    const { rpc, auth } = start();
    const noHeader = await rpc({
      jsonrpc: "2.0",
      id: 8,
      method: "events/stream",
      params: { name: "new_mail" },
    });
    expect(noHeader.status).toBe(400);
    const unknown = (await (
      await rpc(
        { jsonrpc: "2.0", id: 9, method: "events/stream", params: { name: "nope" } },
        { ...auth, "mcp-name": "nope" },
      )
    ).json()) as { error: { code: number } };
    expect(unknown.error.code).toBe(-32023);
  });

  test("the health endpoint reports the last successful poll", async () => {
    const { rpc, base, eventHeaders } = start();
    await rpc({ jsonrpc: "2.0", id: 10, method: "events/poll", params: { name: "new_mail", arguments: {}, cursor: null } }, eventHeaders);
    const health = (await (await fetch(`${base}/health`)).json()) as { lastSuccessfulPoll: string | null };
    expect(health.lastSuccessfulPoll).not.toBeNull();
  });
});
