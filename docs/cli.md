# CLI, daemon and files

One executable. `serve` is the daemon; the other commands configure and install
it. Windows is the target; `serve` also runs on macOS and Linux for development.

## Commands

| Command | What it does |
|---|---|
| `setup` | Asks for the port, the recipient allowlist, whether sending is allowed, and whether Outlook should start at log on. Writes the config and token and prints the MCP client snippet. |
| `doctor` | One line per check, with the fix. Exit code is non-zero when a check fails. |
| `install` | Copies this exe to `%LOCALAPPDATA%\msoutlook-mcp\bin`, registers the daemon Task Scheduler logon task, and starts it. No admin rights. |
| `uninstall` | Stops and deletes the daemon and Outlook tasks. Keeps config unless `--purge`. |
| `upgrade` | Run the **new** exe as `<new>.exe upgrade`. No download. Stops the task, moves the old exe aside, copies itself, runs `version` on the new copy, restarts, and restores the old exe if the self-check fails. `--force` allows a downgrade. |
| `serve` | Runs the daemon in the foreground. This is what the task starts. |
| `config show` / `config set` | Read or change the config, including the send switch, without a restart. |
| `version`, `--help` | Version and usage. Every command has `--help`. |

`install` refuses to run from `bun`; build the exe first. `upgrade` does the same.

## Endpoints

`serve` listens on `127.0.0.1` only.

- `/mcp`: Streamable HTTP MCP. Requires `Authorization: Bearer <token>`.
- `/health`: no token. Reports `connected`, `waiting_for_outlook` or `error`, the version, and the last successful event poll.

Every request must have a loopback `Host`. A present `Origin` must also be
loopback. This blocks DNS rebinding. A missing token or a foreign origin is
rejected before any JSON-RPC handling.

## Files

| Windows | Elsewhere |
|---|---|
| `%LOCALAPPDATA%\msoutlook-mcp\config.json` | `$XDG_CONFIG_HOME/msoutlook-mcp/config.json` (or `~/.config/...`) |
| `%LOCALAPPDATA%\msoutlook-mcp\token` | same config directory |
| `%LOCALAPPDATA%\msoutlook-mcp\state.json`, `events.json` | `$XDG_STATE_HOME/msoutlook-mcp/...` (or `~/.local/state/...`) |
| `%LOCALAPPDATA%\msoutlook-mcp\logs\server.log` | `$XDG_STATE_HOME/msoutlook-mcp/logs/server.log` |

`MSOUTLOOK_MCP_HOME` overrides both directories with one path. Tests use it.

The token and config are written atomically with mode `0600`; on Windows the ACL
is restricted to the current user with `icacls`. The token is never logged and
never written into `config.json`.

## Config keys

```json
{
  "version": 1,
  "port": 8787,
  "allow_send": false,
  "allowed_recipients": ["example.com", "*@contractor.example.net"],
  "shared_mailboxes": [],
  "auto_launch_outlook": false,
  "start_outlook_at_logon": true,
  "poll_interval_ms": 30000,
  "event_folders": ["inbox"],
  "log_level": "info"
}
```

- `allow_send` is read **at send time**, so `config set allow-send on` takes effect without a restart.
- `allowed_recipients`: exact address, domain, `*@domain`, `*.domain` or `/regex/`. Plain entries are literal. An invalid regex stops the config from loading. Matching is case-insensitive and covers To, Cc and Bcc.
- `auto_launch_outlook` is off by default. The logon task already starts Outlook; letting the daemon launch it can surprise the user with a profile picker.
- `start_outlook_at_logon` is what `setup` asks about and `uninstall` removes.

## Task Scheduler

Both tasks run only while the user is signed in, with the interactive token and
least privilege. That is required: COM automation needs the user's desktop
session, so a Windows service cannot work.

- `MSOutlook MCP`: runs `wscript.exe //B "<install>\start-hidden.vbs"` at log on, 15 s delay, one instance, restart on failure, no time limit. The launcher starts `<installed>.exe serve` with window style 0, so no console window appears.
- `MSOutlook MCP Outlook`: runs `cmd /c start "" /min "<OUTLOOK.EXE>"` at log on, 5 s delay. The path comes from the App Paths registry key, not a hardcoded Office folder.

**Minimize Outlook; don't close it.** Closing it stops the connection. To keep it
out of the way, right-click the Outlook icon in the notification area and enable
**Hide When Minimized** (verified against Microsoft Support).

### Why not `--windows-hide-console`

Bun has the flag, and it was checked: it changes the exe to the Windows GUI
subsystem. A GUI-subsystem exe does not inherit the console when started from a
prompt, so `setup`, `doctor` and `version` would print nothing. `install` keeps
the console subsystem and uses the hidden VBS launcher for the task instead.
The launcher is unverified on Windows; see `docs/windows-checklist.md`.
