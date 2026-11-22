import { splitRecipients } from "./util.ts";
import {
  arrayProp,
  boolean,
  booleanProp,
  integer,
  numberProp,
  optionalString,
  requireString,
  schemaObject,
  stringList,
  stringProp,
} from "./args.ts";
import { checkRecipients, confirmRequired, fromMailboxAllowed, recipientValues, sendAllowed } from "./guardrails.ts";
import { fail, ok, type ToolContext, ToolArgumentError, tool } from "./registry.ts";

export function mailWriteTools(ctx: ToolContext) {
  return [
    tool(
      "create_draft",
      "Create a draft in Outlook's Drafts folder. Never sends. Recipients are checked against the allowlist at creation and again at send time.",
      schemaObject(
        {
          to: stringProp("recipient(s), separated by ',' or ';'"),
          subject: stringProp("mail subject"),
          body: stringProp("body, plain text or HTML"),
          cc: stringProp("optional Cc recipients"),
          bcc: stringProp("optional Bcc recipients"),
          html: booleanProp("interpret body as HTML", { default: false }),
          attachments: arrayProp("absolute paths to local files"),
          from_mailbox: stringProp("shared mailbox address or DisplayName (must be in shared_mailboxes)"),
        },
        ["to", "subject", "body"],
      ),
      async (args) => {
        const to = splitRecipients(requireString(args, "to"));
        const cc = splitRecipients(optionalString(args, "cc"));
        const bcc = splitRecipients(optionalString(args, "bcc"));
        const subject = requireString(args, "subject");
        const body = requireString(args, "body");
        const html = boolean(args, "html", false);
        const attachments = stringList(args, "attachments");
        const fromMailbox = optionalString(args, "from_mailbox");

        const blocked =
          checkRecipients(ctx.guardrails, [...to, ...cc, ...bcc]) ??
          fromMailboxAllowed(ctx.guardrails, fromMailbox);
        if (blocked) return blocked;

        const draft = await ctx.bridge.createDraft({
          to,
          cc,
          bcc,
          subject,
          body,
          html,
          attachments,
          from_mailbox: fromMailbox ?? null,
        });
        ctx.logger.info("create_draft", { to_count: to.length, cc_count: cc.length, bcc_count: bcc.length });
        return ok({
          ok: true,
          entry_id: draft.entry_id,
          to,
          cc,
          bcc,
          subject,
          from_mailbox: fromMailbox ?? null,
          drafts_folder: draft.drafts_folder,
          attachments: draft.attachments,
          next_step: "Review the draft in Outlook, then call send_draft(entry_id, confirm=true) or click Send.",
        });
      },
    ),

    tool(
      "send_draft",
      "Send an existing draft. Needs confirm=true and sending enabled. Every recipient is re-checked against the allowlist at send time.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the draft"),
          confirm: booleanProp("must be true to actually send"),
        },
        ["entry_id"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        if (args.confirm !== true) return confirmRequired("send_draft");
        const blockedSend = sendAllowed(ctx.guardrails);
        if (blockedSend) return blockedSend;

        const recipients = await ctx.bridge.getRecipients({ entryId });
        const values = recipientValues(recipients);
        const blocked = checkRecipients(ctx.guardrails, values, true);
        if (blocked) return blocked;
        await ctx.bridge.sendDraft({ entryId });
        ctx.logger.info("send_draft", { recipients: values.length });
        return ok({ ok: true, sent_entry_id: entryId, recipients: values });
      },
    ),

    tool(
      "reply_mail",
      "Reply to a mail. Defaults to saving a draft; direct sending needs confirm=true and sending enabled.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the original mail"),
          body: stringProp("reply body, plain text or HTML"),
          reply_all: booleanProp("reply to all recipients", { default: false }),
          html: booleanProp("interpret body as HTML", { default: false }),
          save_only: booleanProp("save a draft instead of sending (default true)", { default: true }),
          from_mailbox: stringProp("shared mailbox address or DisplayName (forces save_only)"),
          confirm: booleanProp("required when save_only is false"),
        },
        ["entry_id", "body"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        const body = requireString(args, "body");
        const replyAll = boolean(args, "reply_all", false);
        const html = boolean(args, "html", false);
        const saveOnly = boolean(args, "save_only", true);
        const fromMailbox = optionalString(args, "from_mailbox");

        const blockedMailbox = fromMailboxAllowed(ctx.guardrails, fromMailbox);
        if (blockedMailbox) return blockedMailbox;
        // Do not create a stray draft just because a direct send was not confirmed.
        if (!saveOnly && !fromMailbox && args.confirm !== true) return confirmRequired("reply_mail");

        const draft = await ctx.bridge.prepareReply({
          entry_id: entryId,
          body,
          reply_all: replyAll,
          html,
          from_mailbox: fromMailbox ?? null,
        });
        const values = recipientValues(draft.recipients);
        const blocked = checkRecipients(ctx.guardrails, values, false);
        const outside = blocked ? ((blocked.data as { rejected?: string[] }).rejected ?? []) : [];

        if (fromMailbox && !saveOnly) {
          return fail("FROM_MAILBOX_REQUIRES_SAVE_ONLY", "Direct send via reply_mail is disabled when from_mailbox is set; the draft is saved.", {
            saved_draft_entry_id: draft.entry_id,
            drafts_folder: draft.drafts_folder,
            recipients: values,
            recipients_outside_allowlist: outside,
            hint: "Review the shared draft, then click Send in Outlook or call send_draft(entry_id, confirm=true).",
          });
        }

        if (saveOnly) {
          ctx.logger.info("reply_mail_draft", { reply_all: replyAll, recipients: values.length });
          return ok({
            ok: true,
            draft_entry_id: draft.entry_id,
            from_mailbox: fromMailbox ?? null,
            drafts_folder: draft.drafts_folder,
            recipients: values,
            recipients_outside_allowlist: outside,
            next_step: "Review the draft, then call send_draft(draft_entry_id, confirm=true) or click Send.",
          });
        }

        if (args.confirm !== true) return confirmRequired("reply_mail");
        const blockedSend = sendAllowed(ctx.guardrails);
        if (blockedSend) {
          return fail("ALLOW_SEND_DISABLED", "Sending is off; the reply is saved as a draft.", {
            saved_draft_entry_id: draft.entry_id,
            drafts_folder: draft.drafts_folder,
          });
        }
        if (blocked) {
          return fail("RECIPIENTS_NOT_ALLOWED", undefined, {
            rejected: outside,
            saved_draft_entry_id: draft.entry_id,
            hint: "The reply was saved as a draft. Fix the recipients or the allowlist.",
          });
        }
        // Re-read just before Send so a concurrent hand edit is covered too.
        const current = recipientValues(await ctx.bridge.getRecipients({ entryId: draft.entry_id }));
        const blockedAtSend = checkRecipients(ctx.guardrails, current, true);
        if (blockedAtSend) {
          const detail = blockedAtSend.data as {
            message?: string;
            rejected?: string[];
            allowed_recipients?: string[];
            hint?: string;
          };
          return fail("RECIPIENTS_NOT_ALLOWED_AT_SEND", detail.message, {
            rejected: detail.rejected,
            allowed_recipients: detail.allowed_recipients,
            hint: detail.hint,
            saved_draft_entry_id: draft.entry_id,
          });
        }
        await ctx.bridge.sendDraft({ entryId: draft.entry_id });
        ctx.logger.info("reply_mail_sent", { reply_all: replyAll, recipients: current.length });
        return ok({ ok: true, sent: true, recipients: current });
      },
    ),

    tool(
      "move_mail",
      "Move a mail to another folder. Needs confirm=true.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the mail"),
          target_folder: stringProp("folder name or full path"),
          mailbox: stringProp("optional store DisplayName or SMTP address"),
          confirm: booleanProp("must be true"),
        },
        ["entry_id", "target_folder"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        const targetFolder = requireString(args, "target_folder");
        const mailbox = optionalString(args, "mailbox");
        if (args.confirm !== true) return confirmRequired("move_mail");
        const moved = await ctx.bridge.moveMail({ entryId, targetFolder, mailbox: mailbox ?? null });
        ctx.logger.info("move_mail", {});
        return ok({ ok: true, new_entry_id: moved.new_entry_id, target_folder: moved.target_folder });
      },
    ),

    tool(
      "mark_read",
      "Mark a mail read or unread. Needs confirm=true.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the mail"),
          read: booleanProp("true to mark read, false for unread", { default: true }),
          confirm: booleanProp("must be true"),
        },
        ["entry_id"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        const read = boolean(args, "read", true);
        if (args.confirm !== true) return confirmRequired("mark_read");
        await ctx.bridge.markRead({ entryId, read });
        return ok({ ok: true, entry_id: entryId, is_read: read });
      },
    ),

    tool(
      "flag_mail",
      "Set or clear the follow-up flag. Needs confirm=true.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the mail"),
          flag: booleanProp("true to flag, false to clear", { default: true }),
          confirm: booleanProp("must be true"),
        },
        ["entry_id"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        const flag = boolean(args, "flag", true);
        if (args.confirm !== true) return confirmRequired("flag_mail");
        await ctx.bridge.flagMail({ entryId, flag });
        return ok({ ok: true, entry_id: entryId, flagged: flag });
      },
    ),
  ];
}
