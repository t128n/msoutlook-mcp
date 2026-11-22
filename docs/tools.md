# Tools

27 tools, ported from Astral0/outlook-com-mcp. Names and arguments match the
original unless noted. Read tools return JSON; list and search results carry
`truncated` and `coverage` so a silently capped window is visible.

## Phase 1 — read (available now)

| Tool | Arguments | Notes |
|---|---|---|
| `whoami` | | Accounts, mounted stores, inbox counts. |
| `list_folders` | `max_depth`, `mailbox` | Tree, max depth 6. |
| `list_mail` | `folder`, `limit`, `unread_only`, `since_iso`, `sort`, `cursor`, `mailbox` | Summaries only; no body. `sort` is `newest` (default) or `oldest`; a capped page carries an opaque `next_cursor`. |
| `read_mail` | `entry_id`, `max_body_bytes`, `include_html`, `include_transport_headers`, `attachments`, `max_attachment_bytes`, `max_attachment_total_bytes` | Body truncated by UTF-8 bytes. Optional raw transport headers and per-attachment base64 bytes; see [MIME fidelity](#mime-fidelity). |
| `search_mail` | `query`, `folder`, `scope`, `limit`, `mailbox` | DASL `@SQL=` restriction. |
| `download_attachment` | `entry_id`, `index`, `dest_dir` | Saves a file; does not change the mailbox. |
| `health_check` | | Outlook state, guardrails, log path; never throws when Outlook is closed. |
| `guardrails_status` | | Sending switch, allowlist, from-mailbox list. |
| `list_events` | `days_ahead`, `days_back`, `limit`, `calendar` | Iterates with `IncludeRecurrences`, breaks early; no `Restrict` + recurrences. |
| `read_event` | `entry_id`, `max_body_bytes` | Body and attendee response status. |
| `find_freeslots` | `duration_minutes`, `days_ahead`, `business_hours`, `weekdays_only`, `slot_step_minutes`, `limit` | Your own calendar only. |
| `find_freeslots_multi` | `attendees`, same slot options, `include_self` | Exchange `Recipient.FreeBusy`, plus your calendar when asked. |

## Phase 2 — guarded writes (available now)

| Tool | Guardrail |
|---|---|
| `create_draft` | Creates a draft. To, Cc and Bcc are checked at creation. |
| `send_draft` | Needs `confirm=true`, the send switch, and a fresh allowlist check of every To/Cc/Bcc recipient. |
| `reply_mail` | Saves a draft by default. Direct send needs `confirm=true` and the send switch; the saved draft is kept when a check fails. |
| `move_mail` | Needs `confirm=true`. |
| `mark_read` | Needs `confirm=true`. |
| `flag_mail` | Needs `confirm=true`. |
| `create_event_draft` | Creates a draft; attendees are checked. No invite is sent. |
| `send_event_invites` | Needs `confirm=true`, the send switch, and a fresh attendee check. |
| `respond_meeting` | Needs `confirm=true`. `send_response=true` also needs the send switch. |

`from_mailbox` on `create_draft` and `reply_mail` must be listed in
`shared_mailboxes`; `reply_mail` always saves a draft when it is set.

Sending needs the send switch (`config set allow-send on`). The switch and the
allowlist are read from disk at call time, so flipping them needs no restart.
Every To, Cc and Bcc address is re-checked **at send time**, not only when the
draft was created.

## Phase 3 — rules and contacts (available now)

| Tool | Guardrail |
|---|---|
| `list_rules` | Read-only. |
| `toggle_rule` | Needs `confirm=true`. |
| `create_rule` | Needs `confirm=true`, at least one condition and one action. `forward_to` addresses are checked against the allowlist. |
| `delete_rule` | Needs `confirm=true`. |
| `search_contacts` | Read-only GAL search. |
| `get_contact_details` | Read-only. |

Contact search uses ADODB/LDAP **inside the PowerShell helper**, with
`CreateRecipient` as the fallback when the host is not domain-joined. That is the
TypeScript answer for ADODB/LDAP: no native dependency and no second bridge; the
helper already owns COM. `CreateRecipient` stays as the fallback.

`create_rule` checks `forward_to` (the only outgoing action) against the
allowlist. It does not check `from_addresses` or `sent_to`, which are incoming
filters, not recipients. The original checked all three; this is the documented
difference.

A rule with `forward_to` is a send path, so creating one and enabling an
existing forward rule both need the send switch. With sending off, they are
refused.

## Pagination

`list_mail` accepts `sort: "newest" | "oldest"` (default `newest`) and an opaque
`cursor`. When a page is capped, the payload carries `next_cursor`; pass it back
with the same arguments to continue. The final page has `truncated: false` and
no `next_cursor`.

The cursor is `(ReceivedTime, EntryIDs in that second)`, not an offset, because
Outlook's restrictions are inclusive at the boundary. A block of mail that
shares one `ReceivedTime` therefore drains without duplicates or gaps, even
when a page ends inside the block. `unread_only`, `since_iso` and `mailbox`
compose with it. A cursor used with a different `sort` is rejected; omitting
`sort` adopts the cursor's direction.

`truncated` still means "more matching items exist". `search_mail` keeps its
cap and has no cursor yet.

## MIME fidelity

`read_mail` can return the pieces a MIME builder needs, all opt-in:

- `include_transport_headers: true` adds `transport_headers`, the verbatim
  `PR_TRANSPORT_MESSAGE_HEADERS` (`0x007D001F`) block: `From`, `Date`,
  `Message-ID`, `References`, `In-Reply-To` and the `Received` chain as Outlook
  stored them. It is capped at 64 KiB and `transport_headers_truncated` reports
  a cut. The block is untrusted data and is never written to `server.log`.
- `attachments: "base64"` adds `data_base64` to each attachment, plus
  `content_id`, `mime_type` and `is_inline` in both modes. `is_inline` means
  `content_id` is non-empty, which is what a `multipart/related` builder needs;
  `cid:` references in `include_html` match it.
- `max_attachment_bytes` caps one attachment (default 2 MiB) and
  `max_attachment_total_bytes` caps the sum in one response (default 8 MiB).
  A skipped part keeps an empty `data_base64` and a `skipped_reason` of
  `too_large`, `total_cap` or `unavailable`; nothing is silently dropped.

Base64 grows bytes by 4/3, so the default ceiling is about 2.7 MiB per part and
10.7 MiB for the whole response. Clients should not set a 1 MiB read timeout for
a base64 `read_mail`. Binary content is not logged.

Metadata mode (the default) returns the historical attachment fields plus
`content_id`, `mime_type` and `is_inline`; no bytes.

## Differences from the original

- State-changing tools add `confirm=true` where the original performed the
  change directly. This follows the guardrail decision in #5.
- `find_freeslots_multi` accepts either a comma/semicolon string or an array for
  `attendees`.
- `health_check` adds the daemon state, version and log path.
- Errors are returned as a tool result with `isError: true` and a JSON body
  (`OUTLOOK_UNAVAILABLE` with `retryable` when Outlook is away), so a client
  does not treat a closed Outlook as a protocol failure.

## Outlook gotchas kept

- DASL filters are built only in `src/outlook/dasl.ts`; no Jet-style names.
- `IncludeRecurrences` is never combined with `Items.Restrict`.
- X.500 senders are resolved to SMTP through `PropertyAccessor`
  (`PR_SMTP_ADDRESS`, `0x39FE001E`).
- `EntryID` is the identifier; the Internet Message-ID is used as the
  `new_mail` event id when available.
- `find_freeslots` looks at your own calendar only.
