# msoutlook-mcp

An [MCP](https://modelcontextprotocol.io) server for **classic Outlook on Windows** through COM, with polling-based [MCP Events](https://modelcontextprotocol.io/extensions/events/overview). TypeScript on Bun, shipped as one executable.

It is a port of [Astral0/outlook-com-mcp](https://github.com/Astral0/outlook-com-mcp) (Python, MIT). The tool inventory, the write-guardrail model and the documented Outlook gotchas follow that project. It drives the already-authenticated Outlook desktop client, so it needs no app registration and no Microsoft Graph permissions.

> **Status: v0.1.0 released, awaiting Windows verification.** The COM bridge, CLI/daemon, lifecycle, all 27 tools and `new_mail` are implemented. Real Outlook, Task Scheduler and OpenCode behaviour still need the manual Windows checklist; see the [tracking issue](https://github.com/t128n/msoutlook-mcp/issues/1).

## Idea

- **Same safety model as the original:** drafts first, recipient allowlist re-checked at send time, sending off by default.
- **Events:** a `new_mail` event by polling the Inbox, so an agent can wake when mail arrives. Thin payload, no body, persisted cursor.
- **One executable:** `setup`, `doctor`, `install`, `uninstall`, `upgrade`, `serve`, `config`, `version`.
- **A local daemon, not stdio:** loopback HTTP with a bearer token, so the poller keeps running between client sessions.
- **No addon:** a persistent PowerShell process owns the COM objects. A stuck call is killed and replaced.

## Requirements

- Windows 10 or 11
- Outlook **Classic** (the new Outlook has no COM surface)
- No Node, Python or Bun runtime for the released executable

**Minimize Outlook; don't close it.** Closing Outlook stops the connection. The daemon reconnects when Outlook is back, but tools fail while it is gone. To keep Outlook out of the way, right-click its notification-area icon and enable **Hide When Minimized**.

## Install

### mise

```sh
mise use -g github:t128n/msoutlook-mcp@0.1.0
msoutlook-mcp setup
```

`mise install github:t128n/msoutlook-mcp` installs the latest release. The
Windows x64 asset is the supported build; the macOS and Linux assets are
development builds.

### Direct download

Download `msoutlook-mcp-x86_64-pc-windows-msvc.zip` from the
[releases page](https://github.com/t128n/msoutlook-mcp/releases) and run
`msoutlook-mcp.exe`.

### First run

1. Run `msoutlook-mcp setup`. It writes the config and token, asks for the recipient allowlist and the send switch, and prints an MCP client entry.
2. Run `msoutlook-mcp install`. No admin rights; it copies the exe to `%LOCALAPPDATA%\msoutlook-mcp\bin` and registers a Task Scheduler logon task.
3. Run `msoutlook-mcp doctor` if anything does not work. Every line says what to fix.

Start Outlook, sign in, and minimize it. The daemon waits for Outlook and reconnects on its own.

## Let your agent set it up

Paste this into OpenCode (or another coding agent) on the Windows machine:

> Set up `msoutlook-mcp` for OpenCode on this Windows machine with **classic**
> Outlook, not the new Outlook.
>
> 1. Install the latest release with `mise use -g github:t128n/msoutlook-mcp`
>    (or download `msoutlook-mcp-x86_64-pc-windows-msvc.zip` from
>    https://github.com/t128n/msoutlook-mcp/releases). Confirm
>    `msoutlook-mcp version`.
> 2. Ask me for my recipient allowlist (recommended: my own domain), whether to
>    allow sending (default **no**), and whether to start Outlook minimized at
>    log on (recommended **yes**). Do not enable sending unless I explicitly
>    say yes.
> 3. Run `msoutlook-mcp setup --non-interactive --allow "<list>"` with
>    `--allow-send` only if I agreed, and `--outlook-logon` or
>    `--no-outlook-logon` per my answer. Then run `msoutlook-mcp install` and
>    `msoutlook-mcp doctor`.
> 4. Add a remote MCP server named `outlook` to my global OpenCode config using
>    the V2 `mcp.servers` shape: URL `http://127.0.0.1:8787/mcp`,
>    `"oauth": false`, `"protocol": "auto"`, and the header
>    `Authorization: Bearer <token>`. Read the token from
>    `%LOCALAPPDATA%\msoutlook-mcp\token`. Prefer setting
>    `MSOUTLOOK_MCP_TOKEN` in my user environment and referencing
>    `{env:MSOUTLOOK_MCP_TOKEN}`; keep the token out of any public repo.
> 5. Run `opencode mcp list`, confirm `outlook` is connected, then call `whoami`
>    and `health_check` and `list_mail` with `limit: 3`. Do not send mail,
>    create drafts, or change Outlook data during setup.
> 6. Tell me to **minimize Outlook; do not close it**. If a check fails, report
>    the exact `doctor` line and its fix.

The OpenCode config it should produce:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "servers": {
      "outlook": {
        "type": "remote",
        "url": "http://127.0.0.1:8787/mcp",
        "oauth": false,
        "headers": {
          "Authorization": "Bearer {env:MSOUTLOOK_MCP_TOKEN}"
        },
        "protocol": "auto"
      }
    }
  }
}
```

`setup` also prints a generic `mcpServers` snippet with the token. Use the
OpenCode shape above for OpenCode, or the printed snippet for a client that uses
`mcpServers`.

## Tools

All 27 tools are available: read mail and calendar, `health_check`, `whoami`,
`guardrails_status`, attachments and free/busy; the guarded writes (drafts,
`send_draft`, `reply_mail`, moves, flags, event drafts/invites, meeting
responses); rules; and contacts. See [docs/tools.md](docs/tools.md) for the
inventory, arguments, guardrails and differences from the original.

## Safety

- Sending is **off** by default. Every state-changing call creates a draft or needs `confirm=true`.
- The recipient allowlist is re-checked **at send time**, so a draft edited by hand cannot bypass it.
- The daemon listens on `127.0.0.1` only, requires a bearer token, and rejects a foreign `Host` or `Origin`.
- Mail bodies, subjects and tokens are never logged.

## Documentation

- [docs/cli.md](docs/cli.md): commands, files, config keys, Task Scheduler, endpoints.
- [docs/tools.md](docs/tools.md): the 27 tools, phases and argument differences.
- [docs/events.md](docs/events.md): `new_mail`, cursors, push streams and limits.
- [docs/smoke.md](docs/smoke.md): the manual smoke test against a real test mailbox.
- [docs/spike-com.md](docs/spike-com.md): why the PowerShell bridge, what was verified.
- [docs/windows-checklist.md](docs/windows-checklist.md): the manual test list for a real Windows machine.
- [docs/release.md](docs/release.md): release assets, tagging and mise installs.

## Development

```sh
bun install
bun test
bunx tsc --noEmit
bun run build:windows

# against a running daemon and a real test mailbox:
bun scripts/smoke.ts
```

`serve` runs on macOS and Linux for development; `install`, `uninstall` and `upgrade` are Windows-only. Tests use a fake COM bridge, so no Outlook is needed.

## License

MIT. See `LICENSE` and `NOTICE`.
