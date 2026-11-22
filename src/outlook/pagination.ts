import type { MailCursor, MailSummary } from "../com/types.ts";

export const LIST_SORTS = ["newest", "oldest"] as const;
export type ListSort = (typeof LIST_SORTS)[number];

export interface ListCursorState {
  v: 1;
  /** ReceivedTime of the last item on the page, ISO UTC. */
  at: string;
  /** EntryIDs already delivered in the same second as `at`, the tie-breaker. */
  ids: string[];
  sort: ListSort;
}

/**
 * Opaque continuation cursor for `list_mail`.
 *
 * Outlook sorts by ReceivedTime only and its restrictions are inclusive, so an
 * index would repeat or skip items when several mails share a timestamp. The
 * cursor keeps the boundary timestamp plus the EntryIDs already delivered at
 * that exact timestamp, mirroring the `new_mail` cursor in `events/new-mail.ts`.
 */
export function encodeListCursor(state: ListCursorState): string {
  return Buffer.from(JSON.stringify(state), "utf8").toString("base64url");
}

export function decodeListCursor(value: string): ListCursorState | null {
  try {
    const parsed = JSON.parse(Buffer.from(value, "base64url").toString("utf8")) as ListCursorState;
    if (!parsed || parsed.v !== 1) return null;
    if (typeof parsed.at !== "string" || Number.isNaN(new Date(parsed.at).getTime())) return null;
    if (!Array.isArray(parsed.ids) || parsed.ids.some((id) => typeof id !== "string")) return null;
    if (!LIST_SORTS.includes(parsed.sort)) return null;
    return { v: 1, at: parsed.at, ids: parsed.ids, sort: parsed.sort };
  } catch {
    /* invalid cursor */
  }
  return null;
}

/**
 * Build the cursor for the next page, or null when the last item has no time.
 *
 * `previous` is the cursor that produced this page. When the page still ends in
 * the same second, its already-delivered IDs must be carried forward; otherwise
 * a later page would return them again. `received_at` and the COM boundary both
 * have second precision, so every item in that second belongs in the set.
 */
export function cursorForPage(
  items: MailSummary[],
  sort: ListSort,
  previous?: MailCursor | null,
): string | null {
  const last = items[items.length - 1];
  if (!last?.received_at) return null;
  const at = last.received_at;
  const atSeconds = Math.floor(new Date(at).getTime() / 1000);
  const ids = items
    .filter((item) => item.received_at && Math.floor(new Date(item.received_at).getTime() / 1000) === atSeconds)
    .map((item) => item.entry_id);
  if (previous && Math.floor(new Date(previous.at).getTime() / 1000) === atSeconds) {
    for (const id of previous.ids) {
      if (!ids.includes(id)) ids.push(id);
    }
  }
  return encodeListCursor({ v: 1, at, ids, sort });
}
