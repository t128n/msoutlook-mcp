import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FakeComBridge } from "../src/com/fake.ts";
import { Logger } from "../src/util/logger.ts";
import { HealthMonitor } from "../src/server/health.ts";
import {
  PROTOCOL_VERSION,
  hostOf,
  isLoopbackHost,
  sameBearer,
  startMcpServer,
  type McpServerHandle,
} from "../src/server/mcp.ts";

const handles: McpServerHandle[] = [];
const dirs: string[] = [];
afterEach(() => {
  for (const h of handles.splice(0)) h.stop();
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

function start() {
  const dir = mkdtempSync(join(tmpdir(), "msoutlook-http-"));
  dirs.push(dir);
  const bridge = new FakeComBridge();
  const health = new HealthMonitor({ bridge, version: "0.1.0", cacheMs: 0, timeoutMs: 500 });
  const server = startMcpServer({
    config: { port: 0 },
    token: "test-token",
    health,
    version: "0.1.0",
    logger: new Logger(join(dir, "server.log"), "error"),
  });
  handles.push(server);
  const base = `http://127.0.0.1:${server.port}`;
  const auth = { authorization: "Bearer test-token", "content-type": "application/json" };
  const rpc = async (body: unknown, headers: Record<string, string> = auth) =>
    fetch(`${base}/mcp`, { method: "POST", headers, body: JSON.stringify(body) });
  return { server, base, auth, rpc, bridge, health };
}

describe("HTTP daemon", () => {
  test("health needs no token and reports connected", async () => {
    const { base } = start();
    const response = await fetch(`${base}/health`);
    expect(response.status).toBe(200);
    const body = (await response.json()) as { status: string; version: string };
    expect(body.status).toBe("connected");
    expect(body.version).toBe("0.1.0");
  });

  test("health reports waiting_for_outlook when Outlook is closed", async () => {
    const { base, bridge, health } = start();
    bridge.closeOutlook();
    health.invalidate();
    const body = (await (await fetch(`${base}/health`)).json()) as { status: string; outlook: { retryable: boolean } };
    expect(body.status).toBe("waiting_for_outlook");
    expect(body.outlook.retryable).toBe(true);
  });

  test("MCP requires the bearer token", async () => {
    const { base } = start();
    const response = await fetch(`${base}/mcp`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: "{}",
    });
    expect(response.status).toBe(401);
  });

  test("a wrong token is rejected without a length leak", async () => {
    const { base } = start();
    const response = await fetch(`${base}/mcp`, {
      method: "POST",
      headers: { authorization: "Bearer wrong", "content-type": "application/json" },
      body: "{}",
    });
    expect(response.status).toBe(401);
  });

  test("a foreign Origin is rejected", async () => {
    const { base, auth } = start();
    const response = await fetch(`${base}/mcp`, {
      method: "POST",
      headers: { ...auth, origin: "https://evil.example" },
      body: "{}",
    });
    expect(response.status).toBe(403);
  });

  test("initialize negotiates the protocol and advertises tools", async () => {
    const { rpc } = start();
    const response = await rpc({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: PROTOCOL_VERSION } });
    expect(response.status).toBe(200);
    const body = (await response.json()) as {
      result: { protocolVersion: string; capabilities: { tools: unknown }; resultType: string };
    };
    expect(body.result.protocolVersion).toBe(PROTOCOL_VERSION);
    expect(body.result.capabilities.tools).toBeDefined();
    expect(body.result.resultType).toBe("complete");
  });

  test("server/discover lists the protocol", async () => {
    const { rpc } = start();
    const body = (await (await rpc({ jsonrpc: "2.0", id: 2, method: "server/discover" })).json()) as {
      result: { supportedVersions: string[] };
    };
    expect(body.result.supportedVersions).toContain(PROTOCOL_VERSION);
  });

  test("tools/list is empty until the tool phase", async () => {
    const { rpc } = start();
    const body = (await (await rpc({ jsonrpc: "2.0", id: 3, method: "tools/list" })).json()) as {
      result: { tools: unknown[] };
    };
    expect(body.result.tools).toEqual([]);
  });

  test("unknown methods return -32601", async () => {
    const { rpc } = start();
    const body = (await (await rpc({ jsonrpc: "2.0", id: 4, method: "nope" })).json()) as {
      error: { code: number };
    };
    expect(body.error.code).toBe(-32601);
  });

  test("a mismatched Mcp-Name header is 400", async () => {
    const { rpc } = start();
    const response = await rpc(
      { jsonrpc: "2.0", id: 5, method: "events/poll", params: { name: "new_mail" } },
      { authorization: "Bearer test-token", "content-type": "application/json", "mcp-name": "other" },
    );
    expect(response.status).toBe(400);
  });

  test("notifications are accepted with 202", async () => {
    const { rpc } = start();
    const response = await rpc({ jsonrpc: "2.0", method: "notifications/initialized" });
    expect(response.status).toBe(202);
  });

  test("a parse error is -32700", async () => {
    const { base, auth } = start();
    const response = await fetch(`${base}/mcp`, { method: "POST", headers: auth, body: "{" });
    expect(response.status).toBe(400);
    const body = (await response.json()) as { error: { code: number } };
    expect(body.error.code).toBe(-32700);
  });

  test("the server binds loopback only", async () => {
    const { server } = start();
    expect(server.url.startsWith("http://127.0.0.1:")).toBe(true);
  });
});

describe("request guards", () => {
  test("host parsing accepts only loopback names", () => {
    expect(hostOf("127.0.0.1:8080")).toBe("127.0.0.1");
    expect(hostOf("localhost:8080")).toBe("localhost");
    expect(hostOf("[::1]:8080")).toBe("[::1]");
    expect(isLoopbackHost("127.0.0.1:8787")).toBe(true);
    expect(isLoopbackHost("localhost")).toBe(true);
    expect(isLoopbackHost("evil.example")).toBe(false);
    expect(isLoopbackHost("127.0.0.1.evil.example")).toBe(false);
    expect(isLoopbackHost("")).toBe(false);
  });

  test("token comparison is exact", () => {
    expect(sameBearer("Bearer abc", "abc")).toBe(true);
    expect(sameBearer("Bearer abcd", "abc")).toBe(false);
    expect(sameBearer("bearer abc", "abc")).toBe(false);
    expect(sameBearer(null, "abc")).toBe(false);
  });
});
