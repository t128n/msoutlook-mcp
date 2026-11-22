import type { MailRecipient } from "../com/types.ts";
import type { Guardrails } from "../config/guardrails.ts";
import { fail, type ToolResponse } from "./registry.ts";

export function recipientValues(recipients: MailRecipient[]): string[] {
  const values: string[] = [];
  for (const recipient of recipients) {
    const value = recipient.smtp || recipient.address || recipient.name;
    if (value) values.push(value);
  }
  return values;
}

export function checkRecipients(guardrails: Guardrails, recipients: string[], atSend = false): ToolResponse | null {
  if (recipients.length === 0) return null;
  const result = guardrails.checkRecipients(recipients);
  if (result.ok) return null;
  return fail(atSend ? "RECIPIENTS_NOT_ALLOWED_AT_SEND" : "RECIPIENTS_NOT_ALLOWED", undefined, {
    rejected: result.rejected,
    allowed_recipients: guardrails.status().allowed_recipients,
    hint: "Change the recipients or update the allowlist with: msoutlook-mcp config set allowed-recipients <list>",
  });
}

export function sendAllowed(guardrails: Guardrails): ToolResponse | null {
  const check = guardrails.canSend();
  if (check.ok) return null;
  return fail(check.code, check.message, {
    allowed_recipients: guardrails.status().allowed_recipients,
  });
}

export function fromMailboxAllowed(guardrails: Guardrails, mailbox: string | undefined): ToolResponse | null {
  if (!mailbox) return null;
  if (guardrails.sharedMailboxAllowed(mailbox)) return null;
  return fail("FROM_MAILBOX_NOT_ALLOWED", undefined, {
    from_mailbox: mailbox,
    allowed_send_on_behalf_of: guardrails.status().allowed_send_on_behalf_of,
    hint: "Add the mailbox with: msoutlook-mcp config set shared-mailboxes <list>",
  });
}

export function confirmRequired(tool: string): ToolResponse {
  return fail("CONFIRM_REQUIRED", undefined, {
    hint: `Call ${tool} with confirm=true. It changes Outlook state.`,
  });
}
