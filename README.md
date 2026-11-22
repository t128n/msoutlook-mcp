# msoutlook-mcp

An [MCP](https://modelcontextprotocol.io) server for **classic Outlook on Windows** through COM, with polling-based [MCP Events](https://modelcontextprotocol.io/extensions/events/overview). TypeScript on Bun, shipped as one executable.

It is a port of [Astral0/outlook-com-mcp](https://github.com/Astral0/outlook-com-mcp) (Python, MIT). The tool inventory, the write-guardrail model and the documented Outlook gotchas follow that project. It drives the already-authenticated Outlook desktop client, so it needs no app registration and no Microsoft Graph permissions.

> **Status: early implementation.** The COM bridge is in place. The CLI, tools and `new_mail` event are being ported issue by issue; see the [tracking issue](https://github.com/t128n/msoutlook-mcp/issues/1).

## Idea

- **Same safety model as the original:** drafts first, recipient allowlist re-checked at send time, sending off by default.
- **Events:** a `new_mail` event by polling the Inbox, so an agent can wake when mail arrives.
- **One executable** with the commands `setup`, `doctor`, `install`, `uninstall`, `upgrade` and `serve`. `install` registers a Task Scheduler task that runs at log on, because the COM bridge needs your interactive session.
- **No addon:** a persistent PowerShell process owns the COM objects. A stuck call is killed and replaced.

## Requirements

- Windows 10 or 11
- Outlook **Classic** (the new Outlook has no COM surface)
- No Node, Python or Bun runtime for the released executable

**Minimize Outlook; don't close it.** Closing Outlook stops the connection. The daemon reconnects when Outlook is back, but tools fail while it is gone.

## Documentation

- [docs/spike-com.md](docs/spike-com.md): why the PowerShell bridge, what was verified.
- [docs/windows-checklist.md](docs/windows-checklist.md): the manual test list for a real Windows machine.

## Development

```sh
bun install
bun test
bunx tsc --noEmit
bun run build:windows
```

## License

MIT. See `LICENSE` and `NOTICE`.
