# Manual Windows checklist

Every item here needs Windows with **classic Outlook** signed in. Nothing on
this list is claimed to work until it is checked here. Run the daemon in a
normal, non-elevated session: COM automation does not cross elevation.

## Build and start

- [ ] `bun install`, `bun test`, `bunx tsc --noEmit` pass.
- [ ] `bun run build:windows` produces `dist/msoutlook-mcp.exe`.
- [ ] Copy the exe alone to a clean Windows 10/11 machine with no Node and no Bun; `msoutlook-mcp.exe version` prints a version.
- [ ] `msoutlook-mcp.exe doctor` runs without a Bun runtime.

## CLI, daemon and tasks (#3)

- [ ] `setup` writes `%LOCALAPPDATA%\msoutlook-mcp\config.json` and `token`; only the current user can read them (`icacls`).
- [ ] `setup` chooses a free port and prints an MCP snippet with the right URL and token.
- [ ] `serve` starts with Outlook closed and `/health` answers `waiting_for_outlook`.
- [ ] A request to `/mcp` without the bearer token is `401`.
- [ ] A request with `Origin: https://evil.example` is `403`.
- [ ] A request with a foreign `Host` is `403`.
- [ ] `install` copies the exe, registers `MSOutlook MCP`, starts it, and survives a sign-out/sign-in.
- [ ] `uninstall` stops and removes both tasks; `schtasks /Query /TN "MSOutlook MCP"` says it does not exist.
- [ ] `uninstall` keeps the config; `uninstall --purge` removes config, token, state and logs.
- [ ] `doctor` prints one line per check, warns when Outlook is closed, and exits non-zero when a real check fails.
- [ ] `<new>.exe upgrade` replaces the installed exe and restarts the task. Corrupt the new exe so its `version` self-check fails; the old exe is restored and the task starts again.
- [ ] `upgrade` refuses a downgrade without `--force`.
- [ ] No console window appears at log on (the task runs `wscript.exe` with `start-hidden.vbs`).
- [ ] `msoutlook-mcp.exe version`, `doctor` and `setup` print normally when run from `cmd.exe` (the exe stays console-subsystem).

## COM bridge (#2)

- [ ] `ping` connects through `GetActiveObject("Outlook.Application")`, returns `Version` and the default profile.
- [ ] With Outlook closed, `ping` fails with a retryable "Outlook isn't running" error, and does not launch Outlook.
- [ ] With auto-launch enabled in config, `ping` starts Outlook once and reconnects.
- [ ] `list_mail(limit=5)` returns the five newest Inbox items with sender, subject and received time.
- [ ] `list_mail(since_iso=...)` uses a DASL `@SQL=` filter on `ReceivedTime` and returns the right window.
- [ ] A subject containing `'` (for example `O'Brien`) can be searched; the filter does not break.
- [ ] `create_draft(from_mailbox=...)` applies `SentOnBehalfOfName` (PUTREF semantics) and moves the draft to the shared Drafts folder.
- [ ] An in-tenant sender shown as `/O=EXCHANGELABS/...` resolves to an SMTP address through `PR_SMTP_ADDRESS` (`0x39FE001E`).
- [ ] 100 sequential `list_mail(limit=1)` calls complete; note the cold-start time and the average call time.
- [ ] Open a modal Outlook prompt (for example a profile or security prompt) during a call. The call times out, the helper is replaced, and the next call after closing the prompt works.
- [ ] Close Outlook while the daemon runs. Calls fail with a retryable error. `/health` says `waiting_for_outlook` with a retry time.
- [ ] Start Outlook again. Within 60 seconds `/health` says `connected` and calls work, without restarting the daemon.
- [ ] The log shows `outlook_disconnected`, `outlook_reconnect_scheduled` and `outlook_connected`.
- [ ] With `auto_launch_outlook` off, closing Outlook never makes the daemon start it. Turn it on; the next probe starts Outlook once.

## Mail and calendar reads (#5 phase 1)

- [ ] `list_folders` walks the default store and a shared store if one is mounted.
- [ ] `read_mail` returns body, recipients and attachment metadata; `max_body_bytes` truncates.
- [ ] `read_mail(include_transport_headers=true)` on a reply returns a block that parses as RFC 5322 with `References` and `In-Reply-To`; without the flag the field is absent.
- [ ] `read_mail(attachments="base64")` returns an inline image with a `content_id`, `mime_type`, `is_inline=true`, and bytes identical to `download_attachment`; a part over the cap reports `skipped_reason`.
- [ ] `list_mail(sort="oldest", cursor=...)` drains a window larger than one page, including a block of mail that shares one `ReceivedTime`, with no duplicates or gaps.
- [ ] `search_mail(cursor=...)` drains a capped result set; a cursor replayed with a different query or `list_mail` filter is rejected.
- [ ] `download_attachment` saves a file and does not overwrite an existing one.
- [ ] `search_mail` works for subject, sender and subject+body.
- [ ] `list_mail` and `search_mail` report `truncated` and `coverage` when the limit is hit.
- [ ] `list_events` includes a recurring event inside the window.
- [ ] `find_freeslots` ignores all-day and free events, and respects business hours.

## Writes and guardrails (#5 phase 2)

- [ ] `create_draft` appears in Drafts and does not send.
- [ ] With sending off, `send_draft(confirm=true)` is refused and the draft stays.
- [ ] Turn sending on **without restarting** (`config set allow-send on`); `send_draft(confirm=true)` sends.
- [ ] Edit a draft by hand to add an outside recipient; `send_draft` refuses at send time.
- [ ] `reply_mail(save_only=true)` saves a draft; `save_only=false` is refused while sending is off.
- [ ] `move_mail`, `mark_read`, `flag_mail` require `confirm=true` and change only the named item.
- [ ] `create_event_draft` with attendees sends no invite until `send_event_invites(confirm=true)`.
- [ ] `respond_meeting` records accept/tentative/decline; with `send_response=true` the organizer is notified only when sending is on.

## Smoke test (#5)

- [ ] `bun scripts/smoke.ts` passes read-only against the test mailbox.
- [ ] `SMOKE_TO=... bun scripts/smoke.ts --writes` creates a draft and leaves it unsent.
- [ ] `SMOKE_ALLOW_SEND=1 ... --writes --send` sends only when that flag is present.
- [ ] `bun scripts/smoke.ts --rules` creates and deletes a disabled rule.

## Rules and contacts (#5 phase 3)

- [ ] `list_rules` decodes a move-to-folder rule and a subject rule.
- [ ] `create_rule` with `apply_now` moves matching Inbox mail.
- [ ] `delete_rule(confirm=true)` removes the rule.
- [ ] On a domain-joined PC, `search_contacts` returns several GAL results through ADODB/LDAP.
- [ ] On a non-domain-joined PC, `search_contacts` falls back to `CreateRecipient` and returns at most one result.
- [ ] `get_contact_details` returns manager and address fields for an Exchange user.

## Lifecycle (#6)

- [ ] `setup` asks before registering Outlook at log on.
- [ ] After sign-out and sign-in, the Outlook task starts classic Outlook minimized.
- [ ] The daemon task starts after it and connects.
- [ ] If Outlook is closed, `/health` says `waiting_for_outlook`, tools return a retryable error, and the log shows the disconnect.
- [ ] Start Outlook again; `/health` says `connected` within 60 seconds and tools work.
- [ ] `uninstall` removes both tasks and leaves the config unless `--purge`.

## Events (#4)

- [ ] A new mail produces one `new_mail` event within the poll interval.
- [ ] Restart the daemon; old mail is not replayed and new mail is not skipped.
- [ ] Move a delivered mail to another folder; no second event is produced.
- [ ] The payload has no body text.
- [ ] A blocked COM call does not end the event stream; an error notification is sent and polling resumes.
- [ ] OpenCode wakes on the event (also needs the OpenCode spike in t128n/github-events-mcp#2).
