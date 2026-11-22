import { timingSafeEqual } from "node:crypto";
import type { Logger } from "../util/logger.ts";
import type { ToolRegistry } from "../tools/registry.ts";
import type { HealthMonitor } from "./health.ts";

export const PROTOCOL_VERSION = "2026-07-28";
export const SUPPORTED_VERSIONS = [PROTOCOL_VERSION, "2025-06-18"];
const MAX_BODY_BYTES = 1024 * 1024;

export interface McpServerOptions {
  config: { port: number };
  token: string;
  health: HealthMonitor;
  version: string;
  logger: Logger;
  tools?: ToolRegistry;
  hostname?: string;
}

export interface McpServerHandle {
  port: number;
  url: string;
  stop(): void;
}

type JsonRpcId = string | number | null;

interface JsonRpcRequest {
  jsonrpc?: string;
  id?: JsonRpcId;
  method?: string;
  params?: Record<string, unknown>;
}

const LOOPBACK_HOSTS = new Set(["127.0.0.1", "localhost", "[::1]", "::1"]);

export function hostOf(value: string): string | null {
  if (!value) return null;
  if (value.startsWith("[")) {
    const end = value.indexOf("]");
    return end > 0 ? value.slice(0, end + 1).toLowerCase() : null;
  }
  const host = value.split(":")[0];
  return host ? host.toLowerCase() : null;
}

export function isLoopbackHost(value: string): boolean {
  const host = hostOf(value);
  return host !== null && LOOPBACK_HOSTS.has(host);
}

export function sameBearer(header: string | null, token: string): boolean {
  if (!header) return false;
  const expected = `Bearer ${token}`;
  const left = Buffer.from(header);
  const right = Buffer.from(expected);
  if (left.length !== right.length) return false;
  return timingSafeEqual(left, right);
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      "x-content-type-options": "nosniff",
    },
  });
}

const result = (id: JsonRpcId, value: Record<string, unknown>) => ({
  jsonrpc: "2.0",
  id,
  result: { resultType: "complete", ...value },
});

const error = (id: JsonRpcId, code: number, message: string, data?: unknown) => ({
  jsonrpc: "2.0",
  id,
  error: data === undefined ? { code, message } : { code, message, data },
});

/**
 * Loopback-only Streamable HTTP MCP endpoint with a bearer token.
 * `/health` is unauthenticated on purpose; it says whether Outlook is reachable.
 */
export function startMcpServer(options: McpServerOptions): McpServerHandle {
  const hostname = options.hostname ?? "127.0.0.1";
  const server = Bun.serve({
    hostname,
    port: options.config.port,
    idleTimeout: 0,
    async fetch(req) {
      const url = new URL(req.url);

      if (!isLoopbackHost(req.headers.get("host") ?? "")) {
        return json({ error: "bad host" }, 403);
      }
      const origin = req.headers.get("origin");
      if (origin) {
        try {
          if (!isLoopbackHost(new URL(origin).host)) return json({ error: "bad origin" }, 403);
        } catch {
          return json({ error: "bad origin" }, 403);
        }
      }

      if (url.pathname === "/health") {
        if (req.method !== "GET" && req.method !== "HEAD") return json({ error: "method not allowed" }, 405);
        const snapshot = await options.health.snapshot();
        return json({
          status: snapshot.status,
          version: options.version,
          lastSuccessfulPoll: snapshot.lastSuccessfulPoll,
          outlook: snapshot.outlook,
        });
      }

      if (url.pathname !== "/mcp") return json({ error: "not found" }, 404);
      if (!sameBearer(req.headers.get("authorization"), options.token)) {
        return json({ error: "unauthorized" }, 401);
      }
      if (req.method !== "POST") return json({ error: "method not allowed" }, 405);

      const declared = req.headers.get("content-length");
      if (declared && Number(declared) > MAX_BODY_BYTES) return json({ error: "request too large" }, 413);
      const body = await req.text();
      if (Buffer.byteLength(body, "utf8") > MAX_BODY_BYTES) return json({ error: "request too large" }, 413);

      let message: JsonRpcRequest;
      try {
        message = JSON.parse(body) as JsonRpcRequest;
      } catch {
        return json(error(null, -32700, "Parse error"), 400);
      }
      if (!message || typeof message !== "object") return json(error(null, -32600, "Invalid Request"), 400);

      // Notifications (initialized, cancelled) have no id and need no answer.
      if (message.id === undefined) return new Response(null, { status: 202 });

      const id = message.id;
      const method = message.method ?? "";
      const params = message.params ?? {};

      const nameHeader = req.headers.get("mcp-name");
      if (nameHeader !== null && typeof params.name === "string" && nameHeader !== params.name) {
        return json(error(id, -32020, "HeaderMismatch", { header: "Mcp-Name" }), 400);
      }

      switch (method) {
        case "initialize": {
          const requested = typeof params.protocolVersion === "string" ? params.protocolVersion : PROTOCOL_VERSION;
          const protocolVersion = SUPPORTED_VERSIONS.includes(requested) ? requested : PROTOCOL_VERSION;
          return json(
            result(id, {
              protocolVersion,
              capabilities: { tools: { listChanged: false } },
              serverInfo: { name: "msoutlook-mcp", version: options.version },
            }),
          );
        }
        case "server/discover":
          return json(
            result(id, {
              supportedVersions: SUPPORTED_VERSIONS,
              capabilities: { tools: { listChanged: false } },
              _meta: { "io.modelcontextprotocol/serverInfo": { name: "msoutlook-mcp", version: options.version } },
            }),
          );
        case "ping":
          return json(result(id, {}));
        case "tools/list":
          return json(result(id, { tools: options.tools?.list() ?? [] }));
        case "tools/call": {
          if (!options.tools) return json(error(id, -32601, "Method not found"), 200);
          const name = typeof params.name === "string" ? params.name : "";
          const toolArgs =
            params.arguments && typeof params.arguments === "object" && !Array.isArray(params.arguments)
              ? (params.arguments as Record<string, unknown>)
              : {};
          if (!name) return json(error(id, -32602, "InvalidParams", { reason: "tool name is required" }), 200);
          if (!options.tools.has(name)) {
            return json(error(id, -32602, "InvalidParams", { reason: `unknown tool ${name}` }), 200);
          }
          const response = await options.tools.call(name, toolArgs);
          const data = response.data;
          const payload: Record<string, unknown> = {
            content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
          };
          if (data && typeof data === "object" && !Array.isArray(data)) payload.structuredContent = data;
          if (response.isError) payload.isError = true;
          return json(result(id, payload));
        }
        default:
          return json(error(id, -32601, "Method not found"), 200);
      }
    },
    error(e) {
      options.logger.error("http_error", { message: e instanceof Error ? e.message : String(e) });
      return json({ error: "internal error" }, 500);
    },
  });

  const actualPort = server.port ?? options.config.port;
  options.logger.info("server_started", { port: actualPort });
  return {
    port: actualPort,
    url: `http://127.0.0.1:${actualPort}/mcp`,
    stop: () => server.stop(true),
  };
}
