import {
  boolean,
  booleanProp,
  integer,
  numberProp,
  oneOf,
  optionalString,
  requireString,
  schemaObject,
  stringList,
  stringProp,
} from "./args.ts";
import { checkRecipients, confirmRequired, recipientValues, sendAllowed } from "./guardrails.ts";
import { fail, ok, type ToolContext, ToolArgumentError, tool } from "./registry.ts";

export function calendarWriteTools(ctx: ToolContext) {
  return [
    tool(
      "create_event_draft",
      "Create an event or meeting draft. Attendees are checked against the allowlist. No invites are sent until send_event_invites(confirm=true).",
      schemaObject(
        {
          subject: stringProp("event title"),
          start_iso: stringProp("start date/time, ISO 8601"),
          duration_minutes: numberProp("duration (default 30, max 480)", {
            default: 30,
            minimum: 5,
            maximum: 480,
          }),
          body: stringProp("optional body"),
          location: stringProp("optional location"),
          required_attendees: {
            description: "required attendee addresses, separated by ',' or ';'",
            type: ["string", "array"],
            items: { type: "string" },
          },
          optional_attendees: {
            description: "optional attendee addresses, separated by ',' or ';'",
            type: ["string", "array"],
            items: { type: "string" },
          },
        },
        ["subject", "start_iso"],
      ),
      async (args) => {
        const subject = requireString(args, "subject");
        const startIso = requireString(args, "start_iso");
        if (Number.isNaN(new Date(startIso).getTime())) throw new ToolArgumentError(`invalid start_iso: ${startIso}`);
        const duration = integer(args, "duration_minutes", 30, 5, 480);
        const body = optionalString(args, "body");
        const location = optionalString(args, "location");
        const required = stringList(args, "required_attendees");
        const optional = stringList(args, "optional_attendees");
        const blocked = checkRecipients(ctx.guardrails, [...required, ...optional]);
        if (blocked) return blocked;

        const event = await ctx.bridge.createEventDraft({
          subject,
          start_iso: startIso,
          duration_minutes: duration,
          body: body ?? null,
          location: location ?? null,
          required_attendees: required,
          optional_attendees: optional,
        });
        ctx.logger.info("create_event_draft", { required: required.length, optional: optional.length });
        return ok({
          ok: true,
          entry_id: event.entry_id,
          subject: event.subject,
          start: event.start,
          duration_minutes: event.duration_minutes,
          required: event.required,
          optional: event.optional,
          is_meeting: event.is_meeting,
          next_step: event.is_meeting
            ? "Invites NOT sent. Review in Outlook > Calendar, then call send_event_invites(entry_id, confirm=true)."
            : "Event saved in the calendar (no attendees).",
        });
      },
    ),

    tool(
      "send_event_invites",
      "Send invites for an existing meeting. Needs confirm=true and sending enabled; attendees are re-checked at send time.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the meeting draft"),
          confirm: booleanProp("must be true to send the invites"),
        },
        ["entry_id"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        if (args.confirm !== true) return confirmRequired("send_event_invites");
        const blockedSend = sendAllowed(ctx.guardrails);
        if (blockedSend) return blockedSend;

        const recipients = await ctx.bridge.getEventRecipients({ entryId });
        if (recipients.length === 0) {
          return fail("NOT_A_MEETING", "This event has no attendees to invite.");
        }
        const values = recipientValues(recipients);
        const blocked = checkRecipients(ctx.guardrails, values, true);
        if (blocked) return blocked;
        await ctx.bridge.sendMeetingInvites({ entryId });
        ctx.logger.info("send_event_invites", { recipients: values.length });
        return ok({ ok: true, sent_entry_id: entryId, recipients: values });
      },
    ),

    tool(
      "respond_meeting",
      "Accept, tentatively accept or decline a meeting. Needs confirm=true. Notifying the organizer also needs sending enabled.",
      schemaObject(
        {
          entry_id: stringProp("EntryID of the meeting request or appointment"),
          response: stringProp("accepted (default), tentative or declined", {
            enum: ["accepted", "tentative", "declined"],
            default: "accepted",
          }),
          send_response: booleanProp("notify the organizer", { default: false }),
          comment: stringProp("optional note to the organizer"),
          confirm: booleanProp("must be true"),
        },
        ["entry_id"],
      ),
      async (args) => {
        const entryId = requireString(args, "entry_id");
        const response = oneOf(
          args,
          "response",
          ["accepted", "tentative", "declined"] as const,
          "accepted",
        );
        const sendResponse = boolean(args, "send_response", false);
        const comment = optionalString(args, "comment");
        if (args.confirm !== true) return confirmRequired("respond_meeting");
        if (sendResponse) {
          const blockedSend = sendAllowed(ctx.guardrails);
          if (blockedSend) return blockedSend;
        }
        const result = await ctx.bridge.respondMeeting({
          entryId,
          response,
          sendResponse,
          comment: comment ?? null,
        });
        ctx.logger.info("respond_meeting", { response, sent: sendResponse });
        return ok({ ok: true, response, action: result.action });
      },
    ),
  ];
}
