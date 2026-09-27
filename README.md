# hermes-skill-outlook

A Microsoft 365 **Calendar + To Do** skill that talks directly to Microsoft Graph from a single self-contained Python script (plus a one-time auth-bootstrap helper). Built as an agent skill (the Microsoft counterpart to a `google-workspace` skill) but usable as a plain CLI on its own.

Originally proposed upstream as [NousResearch/hermes-agent#25979](https://github.com/NousResearch/hermes-agent/issues/25979).

## What it does

Usable three ways: as an agent skill, as a plain CLI, and as an **MCP server** (see below).

- **Calendar:** read (single / all / shared calendars), create, update, delete, keyword-search, full event detail — with recurrence (daily/weekly/monthly/yearly), all-day events, and series-master safety warnings.
- **Microsoft To Do:** lists, tasks, subtasks ("steps"), create/update/complete/delete, recurrence, list CRUD.
- **Idempotent writes** via content-based `transactionId` (Graph dedupes retries within 6h).
- **Write-op audit log** (append-only JSON-Lines) for "who changed what when" forensics.
- **Autonomous token refresh** — the script refreshes the access token on expiry; you only authenticate once.

Output is always JSON; errors go to stderr with exit 1. No framework, no build step — just Python + `requests`.

## Install

Requires Python 3.8+.

```bash
git clone https://github.com/ByteSide/hermes-skill-outlook.git
cd hermes-skill-outlook
pip install -r requirements.txt
```

## Setup

### 1. Register an Azure app

In the [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade):

1. **New registration.** For supported account types, pick **"Accounts in any organizational directory and personal Microsoft accounts"** — this matches the default `"tenant": "common"`. (If you instead pick *"Personal Microsoft accounts only"*, set `"tenant": "consumers"` in `config.json`; for a single work/school tenant, set your tenant ID.)
2. Under **Authentication → Advanced settings**, set **"Allow public client flows"** to **Yes** (required for the device-code flow).
3. Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Calendars.ReadWrite.Shared`
   - `Tasks.ReadWrite`
   - `Tasks.ReadWrite.Shared`
   - `offline_access`
4. Copy the **Application (client) ID** from the Overview page.

### 2. Configure

```bash
cp .credentials/config.example.json .credentials/config.json
# edit .credentials/config.json, paste your client_id
```

### 3. Authenticate (once)

```bash
python3 scripts/setup_auth.py
```

Follow the printed device-code prompt (open the URL, enter the code, sign in). On success it writes `.credentials/credentials.json`. From then on the token refreshes itself.

### 4. Verify

```bash
python3 scripts/outlook.py status
python3 scripts/outlook.py calendars
```

Optionally set `default_calendar` (the ID from `calendars`) and `default_calendar_name` in `config.json` so bare `events`/`create-event` calls target your preferred calendar.

## Usage

See [`SKILL.md`](SKILL.md) for the full subcommand reference and agent-facing usage notes. Quick taste:

```bash
python3 scripts/outlook.py events --all-calendars --days 7
python3 scripts/outlook.py create-event --subject "Dentist" \
  --start "2026-05-20T10:00:00" --end "2026-05-20T11:00:00" --dry-run
python3 scripts/outlook.py tasks
python3 scripts/outlook.py create-task --title "Bananas" --list "Shopping"
```

## MCP server

The same calendar and To Do commands are also available as a **Model Context
Protocol server** (`scripts/mcp_server.py`, stdio), so any MCP client (Claude
Code, Claude Desktop, other agents) can use them as tools. It is a thin adapter
over the CLI: every tool runs the existing subcommand, so validation, the audit
log, the token refresh and the deduplicated event creation are exactly the CLI's.

Requires Python 3.10+ and the official MCP SDK:

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements-mcp.txt
```

Auth and config are shared with the CLI (steps 1–3 above). Then register it, e.g.
in Claude Code:

```bash
# read-only (default)
claude mcp add m365 -- /path/to/.venv/bin/python /path/to/scripts/mcp_server.py
# allow creating/updating/completing
claude mcp add m365 -e M365_MCP_WRITE=1 -- /path/to/.venv/bin/python /path/to/scripts/mcp_server.py
```

**Safe by default.** Without flags the server is read-only. `M365_MCP_WRITE=1`
adds the write tools; deleting additionally needs `M365_MCP_DELETE=1`. Tools
that are not enabled are not registered at all, so a model never sees them.
Every tool carries MCP annotations (`readOnlyHint`, `destructiveHint`,
`idempotentHint`); `create_task` and `add_step` are marked *not* idempotent,
because a retry would create a second one (only events are deduplicated by
Graph). The write tools take `dry_run` like the CLI. Tool arguments always
reach the CLI as values (`--flag=value`, IDs after `--`), so a task titled
`-h` or `--dry-run` stays a title.

| Mode | Tools |
|---|---|
| read (always) | `status`, `list_calendars`, `list_events`, `search_events`, `get_event`, `list_task_lists`, `list_tasks`, `list_steps`, `audit_log` |
| `M365_MCP_WRITE=1` | `create_event`, `update_event`, `create_task`, `update_task`, `complete_task`, `add_step`, `complete_step` |
| + `M365_MCP_DELETE=1` | `delete_event`, `delete_task` |

Failures come back as readable tool errors (the model sees e.g. *"end must be
strictly after start"* or *"calendar 'X' not found. Known: [...]"*), never as a
crashed server. The commands print JSON to stdout, which over stdio is the MCP
channel, so the adapter captures it and serialises calls with a lock.

## Notes

- **Timezone:** defaults to `UTC`. Override with the `OUTLOOK_TZ` env var (e.g. `OUTLOOK_TZ="Europe/Berlin"`). `--start`/`--end` must be timezone-*naive* ISO (no `Z`, no offset) — the TZ is applied via a Graph `Prefer` header.
- **Credentials never leave your machine.** `.credentials/` is gitignored; only `config.example.json` is tracked.
- **Writes are confirm-first by design.** Every mutating subcommand supports `--dry-run`, which returns the request body without side effects — wire your agent to show the dry run and get a yes before the real call.

## Tests

Offline behavior tests (no network, no credentials) for the recurrence,
idempotency, and validation logic:

```bash
python3 tests/test_offline.py
```

MCP server tests (a real MCP client in-process against the server, Microsoft
Graph faked: no network, no credentials; plus one start over real stdio):

```bash
.venv/bin/python -m pytest tests/ -q
```

## License

MIT — see [LICENSE](LICENSE).
