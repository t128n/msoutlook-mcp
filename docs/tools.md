# Tools

27 tools, ported from Astral0/outlook-com-mcp. Names and arguments match the
original unless noted. Read tools return JSON; list and search results carry
`truncated` and `coverage` so a silently capped window is visible.

## Phase 1 — read (available now)

| Tool | Arguments | Notes |
|---|---|---|
| `whoami` | | Accounts, mounted stores, inbox counts. |
| `list_folders` | `max_depth`, `mailbox` | Tree, max depth 6. |
| `list_mail` | `folder`, `limit`, `unread_only`, `since_iso`, `mailbox` | Summaries only; no body. |
| `read_mail` | `entry_id`, `max_body_bytes`, `include_html` | Body truncated by UTF-8 bytes. |
| `search_mail` | `query`, `folder`, `scope`, `limit`, `mailbox` | DASL `@SQL=` restriction. |
| `download_attachment` | `entry_id`, `index`, `dest_dir` | Saves a file; does not change the mailbox. |
| `health_check` | | Outlook state, guardrails, log path; never throws when Outlook is closed. |
| `guardrails_status` | | Sending switch, allowlist, from-mailbox list. |
| `list_events` | `days_ahead`, `days_back`, `limit`, `calendar` | Iterates with `IncludeRecurrences`, breaks early; no `Restrict` + recurrences. |
| `read_event` | `entry_id`, `max_body_bytes` | Body and attendee response status. |
| `find_freeslots` | `duration_minutes`, `days_ahead`, `business_hours`, `weekdays_only`, `slot_step_minutes`, `limit` | Your own calendar only. |
| `find_freeslots_multi` | `attendees`, same slot options, `include_self` | Exchange `Recipient.FreeBusy`, plus your calendar when asked. |

## Phase 2 — guarded writes (next)

`create_draft`, `send_draft`, `reply_mail`, `move_mail`, `mark_read`, `flag_mail`,
`create_event_draft`, `send_event_invites`, `respond_meeting`.

Every state-changing tool either creates a **draft** or needs `confirm=true`.
Sending needs the send switch (`config set allow-send on`) and every To, Cc and
Bcc address is re-checked against the allowlist **at send time**.

## Phase 3 — rules and contacts

`list_rules`, `toggle_rule`, `create_rule`, `delete_rule`, `search_contacts`,
`get_contact_details`. Contact search uses ADODB/LDAP inside the PowerShell
helper, with `CreateRecipient` as the fallback.

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
