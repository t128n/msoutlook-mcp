/**
 * DASL filter construction.
 *
 * Every `@SQL=` filter for `Items.Restrict` is built here. Outlook rejects a
 * mix of Jet-style names (`[ReceivedTime]`) and DASL property URIs, and its date
 * parser is picky, so nothing else may assemble a restriction string.
 *
 * Behaviour follows Astral0/outlook-com-mcp (MIT); see NOTICE.
 */

export const DASL = {
  subject: "urn:schemas:httpmail:subject",
  fromName: "urn:schemas:httpmail:fromname",
  fromEmail: "urn:schemas:httpmail:fromemail",
  received: "urn:schemas:httpmail:datereceived",
  textDescription: "urn:schemas:httpmail:textdescription",
  read: "urn:schemas:httpmail:read",
} as const;

/** Doubles single quotes, the only escape DASL string literals accept. */
export function escapeDasl(value: string): string {
  return value.replace(/\0/g, "").replace(/'/g, "''");
}

/**
 * Outlook compares `datereceived` in local time and accepts a short ISO 8601
 * value without a zone. Passing a UTC `Z` value shifts the window.
 */
export function formatDaslDate(date: Date): string {
  const p = (n: number, width = 2) => String(n).padStart(width, "0");
  return (
    `${p(date.getFullYear(), 4)}-${p(date.getMonth() + 1)}-${p(date.getDate())}` +
    `T${p(date.getHours())}:${p(date.getMinutes())}:${p(date.getSeconds())}`
  );
}

export function parseDate(value: string | Date): Date {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) throw new Error(`invalid date: ${String(value)}`);
  return date;
}

export function receivedSince(value: string | Date): string {
  return `"${DASL.received}" >= '${formatDaslDate(parseDate(value))}'`;
}

export function receivedOnOrAfter(value: string | Date): string {
  return receivedSince(value);
}

export function unreadOnly(): string {
  return `"${DASL.read}" = 0`;
}

export function and(...clauses: (string | null | undefined | false)[]): string {
  return clauses.filter((c): c is string => typeof c === "string" && c.trim() !== "").join(" AND ");
}

export function or(...clauses: (string | null | undefined | false)[]): string {
  const kept = clauses.filter((c): c is string => typeof c === "string" && c.trim() !== "");
  if (kept.length <= 1) return kept[0] ?? "";
  return "(" + kept.join(" OR ") + ")";
}

export function subjectContains(query: string): string {
  return `"${DASL.subject}" LIKE '%${escapeDasl(query)}%'`;
}

export function bodyContains(query: string): string {
  return `"${DASL.textDescription}" LIKE '%${escapeDasl(query)}%'`;
}

export function senderContains(query: string): string {
  return or(
    `"${DASL.fromName}" LIKE '%${escapeDasl(query)}%'`,
    `"${DASL.fromEmail}" LIKE '%${escapeDasl(query)}%'`,
  );
}

export function searchClause(query: string, scope: "subject_body" | "subject" | "sender" | "all"): string {
  switch (scope) {
    case "subject":
      return subjectContains(query);
    case "sender":
      return senderContains(query);
    case "all":
      return or(subjectContains(query), bodyContains(query), senderContains(query));
    case "subject_body":
    default:
      return or(subjectContains(query), bodyContains(query));
  }
}

export function mailFilter(options: { sinceIso?: string | null; unreadOnly?: boolean }): string {
  return and(options.unreadOnly ? unreadOnly() : null, options.sinceIso ? receivedSince(options.sinceIso) : null);
}

/** The argument `Items.Restrict` expects. */
export function restrictExpression(clause: string): string {
  return "@SQL=" + clause;
}
