#!/usr/bin/env bun
/**
 * Manual smoke test against a running daemon and a real test mailbox.
 *
 * It never sends and never changes a rule unless you ask:
 *   --writes   create a draft to SMOKE_TO (kept in Drafts)
 *   --send     also send that draft; requires SMOKE_ALLOW_SEND=1
 *   --rules    create a disabled rule, then delete it
 *
 * Run with:
 *   MSOUTLOOK_MCP_TOKEN=... bun scripts/smoke.ts
 *   SMOKE_TO=you@example.com bun scripts/smoke.ts --writes
 */
import { appPaths } from "../src/config/paths.ts";
import { readToken } from "../src/config/config.ts";

const args = new Set(process.argv.slice(2));
if (args.has("--help") || args.has("-h")) {
  console.log(`Usage: bun scripts/smoke.ts [--writes] [--send] [--rules]

Environment:
  MSOUTLOOK_MCP_URL    default http://127.0.0.1:8787/mcp
  MSOUTLOOK_MCP_TOKEN  default: the token file in the app data folder
  SMOKE_TO             required for --writes
  SMOKE_QUERY          optional search term for search_mail
  SMOKE_CONTACT        optional name or address for search_contacts
  SMOKE_ATTENDEE       optional attendee for find_freeslots_multi
  SMOKE_ALLOW_SEND     must be 1 for --send`);
  process.exit(0);
}

const url = process.env.MSOUTLOOK_MCP_URL ?? "http://127.0.0.1:8787/mcp";
const token =
  process.env.MSOUTLOOK_MCP_TOKEN ??
  (() => {
    try {
      return readToken(appPaths());
    } catch (e) {
      console.error(`No token: ${e instanceof Error ? e.message : String(e)}`);
      process.exit(1);
    }
  })();
const healthUrl = url.replace(/\/mcp$/, "/health");
const headers = {
  authorization: `Bearer ${token}`,
  "content-type": "application/json",
};

let failures = 0;
let nextId = 1;

async function rpc(method: string, params: Record<string, unknown> = {}, name?: string): Promise<unknown> {
  const response = await fetch(url, {
    method: "POST",
    headers: name ? { ...headers, "mcp-name": name } : headers,
    body: JSON.stringify({ jsonrpc: "2.0", id: nextId++, method, params }),
  });
  if (!response.ok) throw new Error(`${method}: HTTP ${response.status}`);
  const body = (await response.json()) as {
    error?: { code: number; message: string };
    result?: unknown;
  };
  if (body.error) throw new Error(`${method}: ${body.error.message} (${body.error.code})`);
  return body.result;
}

async function callTool(name: string, toolArgs: Record<string, unknown> = {}): Promise<Record<string, unknown>> {
  const result = (await rpc("tools/call", { name, arguments: toolArgs }, name)) as {
    content: { text: string }[];
    structuredContent?: Record<string, unknown>;
    isError?: boolean;
  };
  const data = result.structuredContent ?? (JSON.parse(result.content[0]?.text ?? "{}") as Record<string, unknown>);
  if (result.isError) {
    throw new Error(`${name}: ${JSON.stringify(data)}`);
  }
  return data;
}

async function step(name: string, fn: () => Promise<unknown>): Promise<unknown> {
  try {
    const value = await fn();
    console.log(`ok    ${name}`);
    return value;
  } catch (e) {
    failures += 1;
    console.log(`FAIL  ${name}: ${e instanceof Error ? e.message : String(e)}`);
    return undefined;
  }
}

console.log(`msoutlook-mcp smoke: ${url}`);
const health = await step("health", async () => {
  const response = await fetch(healthUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return response.json();
});
console.log(`      ${JSON.stringify(health)}`);

await step("tools/list", async () => {
  const result = (await rpc("tools/list")) as { tools: { name: string }[] };
  if (result.tools.length !== 27) throw new Error(`expected 27 tools, got ${result.tools.length}`);
});

const who = (await step("whoami", () => callTool("whoami"))) as
  | { inbox_unread?: number; default_store?: string; accounts?: { smtp?: string }[] }
  | undefined;
console.log(`      ${who?.default_store ?? "?"}, unread ${who?.inbox_unread ?? "?"}`);

await step("list_folders", async () => {
  const folders = await callTool("list_folders", { max_depth: 2 });
  if (!Array.isArray(folders) || folders.length === 0) throw new Error("no folders");
});

const mail = (await step("list_mail", () => callTool("list_mail", { limit: 1 }))) as
  | { items?: { entry_id: string }[] }
  | undefined;
const first = mail?.items?.[0];
if (first) {
  await step("read_mail", () => callTool("read_mail", { entry_id: first.entry_id, max_body_bytes: 2000 }));
}

if (process.env.SMOKE_QUERY) {
  await step("search_mail", () => callTool("search_mail", { query: process.env.SMOKE_QUERY!, limit: 3 }));
}

const events = (await step("list_events", () => callTool("list_events", { days_ahead: 2, limit: 3 }))) as
  | { events?: { entry_id: string }[] }
  | undefined;
if (events?.events?.[0]) {
  await step("read_event", () => callTool("read_event", { entry_id: events.events![0]!.entry_id }));
}

await step("find_freeslots", () =>
  callTool("find_freeslots", { duration_minutes: 30, days_ahead: 2, limit: 3 }),
);
if (process.env.SMOKE_ATTENDEE) {
  await step("find_freeslots_multi", () =>
    callTool("find_freeslots_multi", {
      attendees: process.env.SMOKE_ATTENDEE,
      duration_minutes: 30,
      days_ahead: 2,
      limit: 3,
    }),
  );
}

await step("guardrails_status", async () => {
  const status = await callTool("guardrails_status");
  if (typeof status.allow_send_direct !== "boolean") throw new Error("missing allow_send_direct");
});

await step("list_rules", () => callTool("list_rules"));

if (process.env.SMOKE_CONTACT) {
  await step("search_contacts", () => callTool("search_contacts", { query: process.env.SMOKE_CONTACT!, limit: 5 }));
}

if (args.has("--rules")) {
  const name = `msoutlook-mcp smoke ${new Date().toISOString()}`;
  await step("create_rule", () =>
    callTool("create_rule", {
      name,
      subject_contains: ["msoutlook-mcp-smoke"],
      assign_categories: ["msoutlook-mcp-smoke"],
      enabled: false,
      stop_processing: true,
      confirm: true,
    }),
  );
  await step("delete_rule", () => callTool("delete_rule", { name, confirm: true }));
}

if (args.has("--writes")) {
  const to = process.env.SMOKE_TO;
  if (!to) {
    console.log("FAIL  --writes needs SMOKE_TO");
    failures += 1;
  } else {
    const draft = (await step("create_draft", () =>
      callTool("create_draft", {
        to,
        subject: "msoutlook-mcp smoke test",
        body: "This draft was created by the msoutlook-mcp smoke test. Delete it if not needed.",
      }),
    )) as { entry_id?: string } | undefined;
    if (args.has("--send")) {
      if (process.env.SMOKE_ALLOW_SEND !== "1") {
        console.log("FAIL  --send needs SMOKE_ALLOW_SEND=1");
        failures += 1;
      } else if (draft?.entry_id) {
        await step("send_draft", () => callTool("send_draft", { entry_id: draft.entry_id, confirm: true }));
      }
    }
  }
}

console.log(failures === 0 ? "\nall smoke checks passed" : `\n${failures} smoke check(s) failed`);
process.exit(failures === 0 ? 0 : 1);
