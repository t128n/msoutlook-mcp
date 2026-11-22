import type { ComBridge } from "../com/bridge.ts";
import { receivedSince } from "../outlook/dasl.ts";
import {
  defineEvent,
  type EventDefinition,
  type EventOccurrence,
  type PollOptions,
  type PollResult,
} from "../../packages/events/src/index.ts";

export interface NewMailArgs {
  folder?: string;
  unreadOnly?: boolean;
  fromDomains?: string[];
}

export interface NewMailPayload {
  entryId: string;
  folder: string;
  senderSmtp: string | null;
  subject: string;
  receivedAt: string;
  hasAttachments: boolean;
  unread: boolean;
}

export interface NewMailDeps {
  bridge: ComBridge;
  intervalMs?: number;
  now?: () => Date;
}

interface CursorState {
  v: 1;
  /** ReceivedTime of the last delivered item, ISO. */
  at: string;
  /** Event ids delivered at exactly `at`, the tie-breaker. */
  ids: string[];
}

export function encodeCursor(state: CursorState): string {
  return Buffer.from(JSON.stringify(state), "utf8").toString("base64url");
}

export function decodeCursor(value: string): CursorState | null {
  try {
    const parsed = JSON.parse(Buffer.from(value, "base64url").toString("utf8")) as CursorState;
    if (parsed && parsed.v === 1 && typeof parsed.at === "string" && Array.isArray(parsed.ids)) {
      if (!Number.isNaN(new Date(parsed.at).getTime())) return parsed;
    }
  } catch {
    /* invalid cursor */
  }
  return null;
}

function eventIdOf(item: { internet_message_id?: string | null; entry_id: string }): string {
  return item.internet_message_id || item.entry_id;
}

function domainOf(address: string | null): string {
  if (!address || !address.includes("@")) return "";
  return address.slice(address.lastIndexOf("@") + 1).toLowerCase();
}

function normalizeArgs(args: NewMailArgs): Required<NewMailArgs> {
  return {
    folder: args.folder?.trim() || "inbox",
    unreadOnly: args.unreadOnly === true,
    fromDomains: (args.fromDomains ?? []).map((d) => d.trim().toLowerCase()).filter(Boolean),
  };
}

/**
 * `new_mail` by polling the watched folder.
 *
 * First poll with a null cursor starts from now. The cursor is the last
 * ReceivedTime plus the event ids at that exact time, which makes the boundary
 * stable when several mails share a timestamp. The payload is deliberately
 * thin: no body, because mail text is untrusted and `read_mail` exists.
 */
export function createNewMailEvent(deps: NewMailDeps): EventDefinition<NewMailArgs, NewMailPayload> {
  const intervalMs = deps.intervalMs ?? 30_000;
  const now = deps.now ?? (() => new Date());

  return defineEvent<NewMailArgs, NewMailPayload>({
    name: "new_mail",
    description:
      "Fires when mail arrives in the watched Outlook folder. The payload is thin and contains no body; use read_mail with entryId for the text.",
    delivery: ["poll", "push"],
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "folder to watch (default inbox)",
        },
        unreadOnly: {
          type: "boolean",
          description: "only emit unread mail",
          default: false,
        },
        fromDomains: {
          type: "array",
          items: { type: "string" },
          description: "only emit senders in these domains",
        },
      },
      required: [],
    },
    payloadSchema: {
      type: "object",
      properties: {
        entryId: { type: "string" },
        folder: { type: "string" },
        senderSmtp: { type: ["string", "null"] },
        subject: { type: "string" },
        receivedAt: { type: "string", format: "date-time" },
        hasAttachments: { type: "boolean" },
        unread: { type: "boolean" },
      },
      required: ["entryId", "folder", "senderSmtp", "subject", "receivedAt", "hasAttachments", "unread"],
    },

    async poll(cursor: string | null, rawArgs: NewMailArgs, options: PollOptions): Promise<PollResult<NewMailPayload>> {
      const args = normalizeArgs(rawArgs ?? {});
      let truncated = false;
      let state: CursorState;
      if (cursor === null || cursor === undefined) {
        state = { v: 1, at: now().toISOString(), ids: [] };
        return {
          events: [],
          cursor: encodeCursor(state),
          hasMore: false,
          nextPollMs: intervalMs,
          truncated: false,
        };
      }
      const decoded = decodeCursor(cursor);
      if (!decoded) {
        state = { v: 1, at: now().toISOString(), ids: [] };
        truncated = true;
        return {
          events: [],
          cursor: encodeCursor(state),
          hasMore: false,
          nextPollMs: intervalMs,
          truncated,
        };
      }
      state = decoded;

      const limit = Math.max(1, Math.min(options.limit, 50));
      const queryStarted = now();
      const cursorMs = new Date(state.at).getTime();
      const filter = receivedSince(state.at);
      const result = await deps.bridge.listMail({
        folder: args.folder,
        limit: limit + 1,
        filter,
        unread_only: args.unreadOnly,
        since_iso: state.at,
        sort_ascending: true,
      });

      const raw = result.items;
      const fromDomains = new Set(args.fromDomains);
      const matching = raw.filter((item) => {
        if (!item.received_at) return false;
        const receivedMs = new Date(item.received_at).getTime();
        if (receivedMs < cursorMs) return false;
        if (receivedMs === cursorMs && state.ids.includes(eventIdOf(item))) return false;
        if (fromDomains.size > 0 && !fromDomains.has(domainOf(item.sender_smtp))) return false;
        return true;
      });

      const page = matching.slice(0, limit);
      const events: EventOccurrence<NewMailPayload>[] = page.map((item) => ({
        eventId: eventIdOf(item),
        name: "new_mail",
        timestamp: item.received_at!,
        data: {
          entryId: item.entry_id,
          folder: item.folder ?? args.folder,
          senderSmtp: item.sender_smtp,
          subject: item.subject,
          receivedAt: item.received_at!,
          hasAttachments: item.has_attachments,
          unread: !item.is_read,
        },
      }));

      const hasMore = raw.length > limit || matching.length > limit;
      let nextState: CursorState;
      if (page.length > 0) {
        const last = page[page.length - 1]!;
        const at = last.received_at!;
        const lastMs = new Date(at).getTime();
        const ids = page
          .filter((item) => item.received_at && new Date(item.received_at).getTime() === lastMs)
          .map((item) => eventIdOf(item));
        if (new Date(state.at).getTime() === lastMs) {
          for (const id of state.ids) if (!ids.includes(id)) ids.push(id);
        }
        nextState = { v: 1, at, ids };
      } else {
        const examined = raw[raw.length - 1]?.received_at;
        if (examined) {
          const examinedMs = new Date(examined).getTime();
          nextState = { v: 1, at: examined, ids: examinedMs === cursorMs ? [...state.ids] : [] };
        } else {
          // Nothing in the window. Advance to just before the query so the
          // window does not grow forever, while leaving a 2 s safety margin.
          const safe = new Date(Math.max(cursorMs, queryStarted.getTime() - 2_000));
          nextState = { v: 1, at: safe.toISOString(), ids: [] };
        }
      }

      return {
        events,
        cursor: encodeCursor(nextState),
        hasMore,
        nextPollMs: intervalMs,
        truncated,
      };
    },
  });
}
