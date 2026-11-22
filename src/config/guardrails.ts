import { type AppConfig, loadConfig } from "./config.ts";
import { type AllowlistResult, checkRecipients } from "./matcher.ts";
import type { AppPaths } from "./paths.ts";

/** Guardrail model ported from Astral0/outlook-com-mcp (MIT); see NOTICE. */
export interface GuardrailsStatus {
  allow_send_direct: boolean;
  allowed_recipients: string[];
  allowed_send_on_behalf_of: string[];
  unrestricted: boolean;
  policy: string;
}

export type SendCheck = { ok: true } | { ok: false; code: "ALLOW_SEND_DISABLED"; message: string };

/**
 * The send switch and the allowlist are read from disk at check time, not at
 * startup, so flipping either needs no restart. A draft edited in Outlook is
 * checked again immediately before `Send()`.
 */
export class Guardrails {
  private readonly config: () => AppConfig;

  constructor(config: AppConfig | (() => AppConfig)) {
    this.config = typeof config === "function" ? config : () => config;
  }

  static fromPaths(paths: AppPaths): Guardrails {
    return new Guardrails(() => loadConfig(paths));
  }

  status(): GuardrailsStatus {
    const config = this.config();
    return {
      allow_send_direct: config.allow_send,
      allowed_recipients: [...config.allowed_recipients],
      allowed_send_on_behalf_of: [...config.shared_mailboxes],
      unrestricted: config.allowed_recipients.length === 0,
      policy:
        "create_draft never sends. Sending requires the switch (config allow_send) and confirm=true, " +
        "and every To, Cc and Bcc address is checked against the allowlist at send time. " +
        "An empty allowlist means no restriction.",
    };
  }

  checkRecipients(recipients: string[]): AllowlistResult {
    return checkRecipients(this.config().allowed_recipients, recipients);
  }

  canSend(): SendCheck {
    if (this.config().allow_send) return { ok: true };
    return {
      ok: false,
      code: "ALLOW_SEND_DISABLED",
      message:
        "Sending is disabled. The draft is still saved; review it in Outlook and click Send, " +
        "or enable sending with: msoutlook-mcp config set allow-send on",
    };
  }

  sharedMailboxAllowed(name: string): boolean {
    const wanted = name.toLowerCase();
    return this.config().shared_mailboxes.some((value) => value.toLowerCase() === wanted);
  }
}
