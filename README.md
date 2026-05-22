# hermes-skill-outlook

A Microsoft 365 **Calendar + To Do** skill that talks directly to Microsoft Graph from a single self-contained Python script. Built as an agent skill (the Microsoft counterpart to a `google-workspace` skill) but usable as a plain CLI on its own.

Originally proposed upstream as [NousResearch/hermes-agent#25979](https://github.com/NousResearch/hermes-agent/issues/25979).

## What it does

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
python scripts/setup_auth.py
```

Follow the printed device-code prompt (open the URL, enter the code, sign in). On success it writes `.credentials/credentials.json`. From then on the token refreshes itself.

### 4. Verify

```bash
python scripts/outlook.py status
python scripts/outlook.py calendars
```

Optionally set `default_calendar` (the ID from `calendars`) and `default_calendar_name` in `config.json` so bare `events`/`create-event` calls target your preferred calendar.

## Usage

See [`SKILL.md`](SKILL.md) for the full subcommand reference and agent-facing usage notes. Quick taste:

```bash
python scripts/outlook.py events --all-calendars --days 7
python scripts/outlook.py create-event --subject "Dentist" \
  --start "2026-05-20T10:00:00" --end "2026-05-20T11:00:00" --dry-run
python scripts/outlook.py tasks
python scripts/outlook.py create-task --title "Bananas" --list "Shopping"
```

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

## License

MIT — see [LICENSE](LICENSE).
