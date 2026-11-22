# `new_mail` event

The server supports the MCP Events extension (`io.modelcontextprotocol/events`,
protocol revision `2026-07-28`) with one event type, `new_mail`.

It polls Outlook locally through COM. No webhook, no mailbox subscription, no
Graph. The poll interval is `poll_interval_ms` in the config (default 30 s).

## Subscribe

- `events/list` returns `new_mail` with `delivery: ["poll", "push"]`, its
  `inputSchema` and its `payloadSchema`.
- `events/poll` polls once.
- `events/stream` opens an SSE stream for push delivery.
- Requests for `events/poll` and `events/stream` must carry `Mcp-Name: new_mail`
  (Streamable HTTP routing header). A mismatch is `400 HeaderMismatch`.
- An unknown event name is `-32023 NotFound` with `data.kind = "event"`.

## Arguments

```json
{
  "folder": "inbox",
  "unreadOnly": false,
  "fromDomains": ["example.com"]
}
```

- `folder`: the watched folder, default `inbox`.
- `unreadOnly`: only emit unread mail.
- `fromDomains`: only emit senders in these domains, matched on the resolved
  SMTP address (not the X.500 address).

## Payload

Thin on purpose. **No body.** Mail text is untrusted; use `read_mail` with
`entryId` for the content.

```json
{
  "entryId": "00000000...",
  "folder": "\\\\Personal Mailbox\\Inbox",
  "senderSmtp": "sender@example.com",
  "subject": "Quarterly report",
  "receivedAt": "2026-04-20T09:15:00Z",
  "hasAttachments": true,
  "unread": true
}
```

`eventId` is the Internet Message-ID when Outlook exposes one, otherwise the
`EntryID`. Moving a delivered mail does not produce a second event.

## Cursor

The cursor is opaque to clients. Internally it is the last `ReceivedTime` plus
the event ids at exactly that timestamp, which keeps the boundary stable when
several mails share a timestamp.

- The first poll with `cursor: null` starts **from now** and returns no events.
- The last cursor per subscription (`name` + arguments) is persisted in
  `events.json`. A daemon restart resumes there: old mail is not replayed and
  new mail is not skipped. A client that reconnects with `null` also resumes
  from the stored cursor.
- A cursor that cannot be decoded resets to now and returns `truncated: true`.
- A batch cap returns `hasMore: true`; poll again immediately with the returned
  cursor to drain the backlog. `truncated` is reserved for a real gap.

## Push behaviour

- One poll loop per distinct `(name, arguments)` is shared by every stream that
  asks for it.
- Frames are `notifications/events/active`, `notifications/events/event`,
  `notifications/events/heartbeat` and `notifications/events/error`. Every frame
  carries the stream id in `_meta["io.modelcontextprotocol/subscriptionId"]`.
- A failed poll sends `notifications/events/error` and backs off from 2 s to
  60 s. The stream stays open. When Outlook is back, polling resumes from the
  cursor with no duplicates.
- Limits: 50 events per poll, 64 KiB per payload (larger events are dropped and
  the batch is marked `truncated`), 32 concurrent subscriptions. The last two
  return `-32025 ResourceExhausted`.
- Heartbeats default to 15 s; the requested `heartbeatMs` is clamped to
  1 000–30 000 ms and echoed in `active`.

## Implementation

`packages/events` is a **local vendored** implementation of the agreed authoring
interface (`defineEvent`, `poll` returning `{ events, cursor, hasMore,
nextPollMs, truncated? }`). It is not a dependency on
t128n/github-events-mcp because that package does not exist yet. The wire format
follows the spike in t128n/github-events-mcp#6.

- `src/events/new-mail.ts`: the Outlook poll and cursor.
- `src/events/hub.ts`: shared poll loops, push streams, backoff and limits.
- `src/events/state.ts`: persisted cursors and the last poll time.

## OpenCode

The wire format is tested with a plain `fetch` client. Whether OpenCode wakes on
the event is **not verified**; it depends on the OpenCode spike in
t128n/github-events-mcp#2.
