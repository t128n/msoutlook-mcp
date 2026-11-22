# Smoke test

`scripts/smoke.ts` exercises a **running daemon** against a real test mailbox.
It is not part of CI; it is the manual check from #5.

```sh
# read-only
bun scripts/smoke.ts

# also create a draft (kept in Drafts)
SMOKE_TO=you@example.com bun scripts/smoke.ts --writes

# also send that draft (explicit, off by default)
SMOKE_TO=you@example.com SMOKE_ALLOW_SEND=1 bun scripts/smoke.ts --writes --send

# also create a disabled rule and delete it
bun scripts/smoke.ts --rules
```

The token comes from `MSOUTLOOK_MCP_TOKEN` or the token file written by
`setup`. `MSOUTLOOK_MCP_URL` defaults to `http://127.0.0.1:28471/mcp`.

## What it covers

- Phase 1: `health`, `tools/list` (27), `whoami`, `list_folders`, `list_mail`,
  `read_mail`, `search_mail` (with `SMOKE_QUERY`), `list_events`, `read_event`,
  `find_freeslots`, `find_freeslots_multi` (with `SMOKE_ATTENDEE`),
  `guardrails_status`.
- Phase 2 (only with `--writes`): `create_draft` to `SMOKE_TO`; `send_draft`
  only with `--send` and `SMOKE_ALLOW_SEND=1`. The draft is left in Drafts
  unless sent; the script has no delete tool.
- Phase 3: `list_rules`, `search_contacts` (with `SMOKE_CONTACT`); with
  `--rules`, a **disabled** rule is created and then deleted.

It prints `ok` or `FAIL` per step and exits non-zero on a failure. It never
prints the token and never includes a mail body in its own output.

## Not verified

The real Outlook responses behind every step. Use
`docs/windows-checklist.md` for the full pass.
