# msoutlook-mcp

An [MCP](https://modelcontextprotocol.io) server for **classic Outlook on Windows** through COM, with polling-based [MCP Events](https://modelcontextprotocol.io/extensions/events/overview). TypeScript on Bun, shipped as one executable.

It is a port of [Astral0/outlook-com-mcp](https://github.com/Astral0/outlook-com-mcp) (Python, MIT). It drives the already-authenticated Outlook desktop client, so it needs no app registration and no Microsoft Graph permissions.

> **Status: planning.** Nothing is built yet. See the issues.

## Idea

- **Same safety model as the original:** drafts first, recipient allowlist re-checked at send time, sending off by default.
- **Events:** a `new_mail` event by polling the Inbox, so an agent can wake when mail arrives.
- **One executable** with the commands `setup`, `doctor`, `install`, `uninstall` and `upgrade`. `install` registers a Task Scheduler task that runs at log on, because the COM bridge needs your interactive session.

## Requirements

- Windows 10 or 11
- Outlook **Classic** (the new Outlook has no COM surface)

## License

MIT. See `LICENSE` and `NOTICE`.
