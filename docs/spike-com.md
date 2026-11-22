# Spike: talking to classic Outlook from Bun

Issue: [#2](https://github.com/t128n/msoutlook-mcp/issues/2). Decision already
made: **a persistent Windows PowerShell process**, no native addon.

## Recommendation

Keep option 2. A long-lived `powershell.exe` owns the COM objects and speaks
JSON lines with the Bun process. winax is dropped: a native N-API addon inside a
compiled Bun exe is an unknown, and one PowerShell process costs far less than a
mailbox integration that does not start.

The real helper is `src/com/script.ts`. The host is `src/com/powershell.ts`. The
contract is `ComBridge` in `src/com/bridge.ts`.

## Protocol

```
bun -> helper   {"id":7,"op":"listMail","folder":"inbox","limit":20}\n
helper -> bun   {"id":7,"ok":true,"result":{"items":[...],"truncated":false}}\n
helper -> bun   {"id":7,"ok":false,"error":{"code":-2147023174,"message":"..."}}\n
```

- `powershell.exe -NoLogo -NoProfile -NonInteractive -Sta -ExecutionPolicy Bypass -EncodedCommand <script>`
- Window hidden (`windowsHide`), stdin/stdout are pipes.
- One request per line, one answer per line, matched by `id`, so concurrent calls do not block each other.
- Non-ASCII is written as `\uXXXX`; both directions are pure ASCII, so the console code page is irrelevant.
- Every call has a timeout. A timeout kills the helper; the next call starts a fresh one. A modal Outlook prompt therefore becomes a retryable error, not a hung daemon.
- The helper resets its cached `Outlook.Application` and `MAPI` namespace after **any** error. A closed Outlook reconnects on the next call.
- HRESULTs are mapped in `src/com/errors.ts` to `not_installed`, `not_running`, `connection_lost`, `timeout`, `access_denied` or `com`. Only the connection-shaped ones are retryable.

## What is verified

Run on macOS on 2026-11-22 with Bun 1.4.4 and PowerShell 7.6.6. These prove the
host/protocol wiring and the helper's shape, **not** COM behaviour.

| Check | Result |
|---|---|
| `bun build --compile --target=bun-windows-x64` | Produces `dist/msoutlook-mcp.exe`, `PE32+ executable (console) x86-64, for MS Windows`, 110 MB. Not executed. |
| Helper parses | `Parser::ParseFile` on the extracted script: `PARSE_OK tokens=9722`. |
| Helper loop under real PowerShell | Sent `ping` and an unknown op. It answered two JSON lines, the first a COM error because this host has no `GetActiveObject`, the second `UNKNOWN_OP`. The loop survived the first error. |
| Host protocol test | Spawns `tests/fixtures/com-stub.ts`, which speaks the same protocol. Covers ids, concurrent calls, stdout noise, HRESULT mapping, timeout + fresh process, helper exit, ASCII round-trip. `bun test` passes. |
| Timeout replaces the helper | Tested against the stub: a `hang` rejects as `timeout`, the next `ping` starts a new process. |
| ASCII escaping | Unit test asserts the wire form has no byte above `0x7e` and that `JSON.parse` restores the text. |

## What is not verified

Everything that needs Windows and classic Outlook. None of these are claimed to
work yet; `docs/windows-checklist.md` has the steps.

| Check | Why not here |
|---|---|
| `GetActiveObject("Outlook.Application")`, `Version`, default profile | No Windows, no Outlook. |
| Newest Inbox items with sender, subject, received time | Same. |
| `Items.Restrict(@SQL=...)` on `ReceivedTime` | Same. The DASL builder has unit tests; Outlook's parser does not. |
| A property that takes an object reference (`SentOnBehalfOfName`, `MoveToFolder.Folder`) | Same. |
| X.500 to SMTP through `PropertyAccessor` (`0x39FE001E`) | Same. |
| Compiled exe on a clean Windows box | The exe is built but cannot be run here. |
| Cold start and 100 calls in a row | Not measured. A number from macOS PowerShell would not describe Windows COM. |
| Modal prompt: timeout instead of a hang | The host timeout is tested with a stub. A real Outlook prompt is not. |
| Outlook closed and restarted while the daemon runs | The helper resets after errors and the host restarts a dead helper. End-to-end Outlook restart belongs to #6. |
| Constrained-language mode | `doctor` will detect it; not tested. |

## Costs and limits

- One extra process per server, serialised calls, and the helper script embedded in the exe.
- The helper must be STA. `-Sta` is supported by Windows PowerShell 5.1 and rejected by PowerShell 7 on non-Windows; the exe always uses `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`.
- PowerShell in constrained-language mode cannot create the COM object. `doctor` reports it.
- Sending is still a TypeScript guardrail. The helper only performs the operation after the tool layer has checked the allowlist and the send switch.
