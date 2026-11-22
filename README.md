# msoutlook-mcp

An [MCP](https://modelcontextprotocol.io) server for **classic Outlook on Windows** through COM, with polling-based [MCP Events](https://modelcontextprotocol.io/extensions/events/overview). TypeScript on Bun, shipped as one executable.

It is a port of [Astral0/outlook-com-mcp](https://github.com/Astral0/outlook-com-mcp) (Python, MIT). The tool inventory, the write-guardrail model and the documented Outlook gotchas follow that project. It drives the already-authenticated Outlook desktop client, so it needs no app registration and no Microsoft Graph permissions.

> **Status: early implementation.** The COM bridge, CLI and daemon are in place. The tools and `new_mail` are being ported issue by issue; see the [tracking issue](https://github.com/t128n/msoutlook-mcp/issues/1).

## Idea

- **Same safety model as the original:** drafts first, recipient allowlist re-checked at send time, sending off by default.
- **Events:** a `new_mail` event by polling the Inbox, so an agent can wake when mail arrives.
- **One executable:** `setup`, `doctor`, `install`, `uninstall`, `upgrade`, `serve`, `config`, `version`.
- **A local daemon, not stdio:** loopback HTTP with a bearer token, so the poller keeps running between client sessions.
- **No addon:** a persistent PowerShell process owns the COM objects. A stuck call is killed and replaced.

## Requirements

- Windows 10 or 11
- Outlook **Classic** (the new Outlook has no COM surface)
- No Node, Python or Bun runtime for the released executable

**Minimize Outlook; don't close it.** Closing Outlook stops the connection. The daemon reconnects when Outlook is back, but tools fail while it is gone. To keep Outlook out of the way, right-click its notification-area icon and enable **Hide When Minimized**.

## Install

1. Run `msoutlook-mcp.exe setup`. It writes the config and token, asks for the recipient allowlist and the send switch, and prints an MCP client entry.
2. Run `msoutlook-mcp.exe install`. No admin rights; it copies the exe to `%LOCALAPPDATA%\msoutlook-mcp\bin` and registers a Task Scheduler logon task.
3. Run `msoutlook-mcp.exe doctor` if anything does not work. Every line says what to fix.

Start Outlook, sign in, and minimize it. The daemon waits for Outlook and reconnects on its own.

## Tools

Phase 1 is available: read mail and calendar, `health_check`, `whoami`,
`guardrails_status`, attachments and free/busy. The guarded writes follow, then
rules and contacts. See [docs/tools.md](docs/tools.md) for the inventory,
arguments and differences from the original.

## Safety

- Sending is **off** by default. Every state-changing call creates a draft or needs `confirm=true`.
- The recipient allowlist is re-checked **at send time**, so a draft edited by hand cannot bypass it.
- The daemon listens on `127.0.0.1` only, requires a bearer token, and rejects a foreign `Host` or `Origin`.
- Mail bodies, subjects and tokens are never logged.

## Documentation

- [docs/cli.md](docs/cli.md): commands, files, config keys, Task Scheduler, endpoints.
- [docs/tools.md](docs/tools.md): the 27 tools, phases and argument differences.
- [docs/spike-com.md](docs/spike-com.md): why the PowerShell bridge, what was verified.
- [docs/windows-checklist.md](docs/windows-checklist.md): the manual test list for a real Windows machine.

## Development

```sh
bun install
bun test
bunx tsc --noEmit
bun run build:windows
```

`serve` runs on macOS and Linux for development; `install`, `uninstall` and `upgrade` are Windows-only. Tests use a fake COM bridge, so no Outlook is needed.

## License

MIT. See `LICENSE` and `NOTICE`.
