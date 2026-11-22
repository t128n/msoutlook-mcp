import { errorText } from "../com/errors.ts";
import type { PingResult } from "../com/types.ts";
import { schemaObject } from "./args.ts";
import { ok, type ToolContext, tool } from "./registry.ts";

export function healthTools(ctx: ToolContext) {
  return [
    tool(
      "health_check",
      "Server diagnostic: Outlook reachable, COM latency, active guardrails and log path. Call this first when something is broken.",
      schemaObject({}, []),
      async () => {
        const started = Date.now();
        const guardrails = ctx.guardrails.status();
        let ping: PingResult | undefined;
        let failure: string | undefined;
        try {
          ping = await ctx.bridge.ping({ timeoutMs: 5_000 });
        } catch (e) {
          failure = errorText(e);
        }
        const snapshot = await ctx.health.snapshot();
        const base: Record<string, unknown> = {
          server_pid: process.pid,
          version: ctx.version,
          runtime: process.versions.bun ? `bun ${process.versions.bun}` : process.version,
          platform: `${process.platform} ${process.arch}`,
          log_path: ctx.paths.logFile,
          guardrails,
          outlook_running: Boolean(ping),
          connect_ms: Date.now() - started,
          status: ping ? "OK" : "ERROR",
          health: {
            state: snapshot.status,
            last_success_at: snapshot.lastSuccessfulCall,
            last_poll_at: snapshot.lastSuccessfulPoll,
          },
        };
        if (ping) {
          base.outlook_version = ping.version;
          base.default_store = ping.default_store;
          base.accounts = ping.accounts;
          base.inbox = { count: ping.inbox_count, unread: ping.inbox_unread };
          base.calendar = { count: ping.calendar_count };
        } else {
          base.error = failure;
          base.hint =
            "Check that classic Outlook is running. Minimize Outlook instead of closing it; the daemon reconnects on its own.";
        }
        return ok(base);
      },
    ),

    tool(
      "whoami",
      "Current Outlook identity, accounts and mounted stores. It can be used as a COM smoke test.",
      schemaObject({}, []),
      async () => {
        const ping = await ctx.bridge.ping({ timeoutMs: 5_000 });
        const folders = await ctx.bridge.listFolders({ maxDepth: 1 });
        const defaultName = ping.default_store;
        const mailboxes = folders.map((folder) => {
          const display = folder.store ?? folder.name ?? "(unknown)";
          const isDefault = Boolean(defaultName && display === defaultName);
          return {
            display_name: display,
            is_default: isDefault,
            from_mailbox_allowed: isDefault || ctx.guardrails.sharedMailboxAllowed(display),
            inbox_unread: folder.unread_count,
          };
        });
        return ok({
          default_store: ping.default_store,
          accounts: ping.accounts,
          accessible_mailboxes: mailboxes,
          allowed_send_on_behalf_of: ctx.guardrails.status().allowed_send_on_behalf_of,
          inbox_count: ping.inbox_count,
          inbox_unread: ping.inbox_unread,
          outlook_version: ping.version,
        });
      },
    ),

    tool(
      "guardrails_status",
      "Current sending switch, recipient allowlist and from-mailbox allowlist. Read-only.",
      schemaObject({}, []),
      async () => ok(ctx.guardrails.status()),
    ),
  ];
}
