---
name: outlook
description: "Microsoft 365 Calendar + Microsoft To Do via Graph API. Read/create/update/complete events and tasks, with recurrence, subtasks, multi-calendar aggregation, keyword search, and a write-op audit log."
version: 1.2.0
metadata:
  hermes:
    tags: [outlook, calendar, tasks, microsoft, graph, productivity]
    related_skills: []
---

# Outlook / Microsoft 365

Talks directly to Microsoft Graph against a personal Microsoft 365 account. Auth runs over a refresh token (in `.credentials/credentials.json`), refreshed autonomously on the first call once the access token has expired.

**What it does:**
- Read calendars — the default calendar is configurable in `config.json`
- Create / update / delete events — confirm with the user before any write
- Microsoft To Do — read, create, complete tasks (writes only after user confirmation)

## Setup (one time)

1. Register an app in [Azure Portal → App registrations](https://portal.azure.com) (personal Microsoft accounts supported). Add the delegated Graph scopes `User.Read`, `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`, `Tasks.ReadWrite`, `Tasks.ReadWrite.Shared`, plus `offline_access`. Enable the "Allow public client flows" toggle (device-code flow).
2. Copy `.credentials/config.example.json` to `.credentials/config.json` and fill in your `client_id`.
3. Run `python scripts/setup_auth.py` and follow the device-code prompt. This writes `.credentials/credentials.json`.
4. Optionally set `default_calendar` / `default_calendar_name` in `config.json` once you know your calendar IDs (run `outlook.py calendars`).

See `README.md` for the full walkthrough.

## Script

Single executable: `scripts/outlook.py`

```bash
python scripts/outlook.py <subcommand> [args]
```

Output is always JSON; errors go to stderr + exit 1.

Optional env var `OUTLOOK_TZ` overrides the timezone (default `Europe/Berlin`).

## When to use which subcommand

| Intent | Subcommand |
|---|---|
| "What's on today / tomorrow / this week?" | `events --all-calendars` (otherwise only the default calendar is seen) |
| "Which calendars do I have?" | `calendars` |
| "When was X again?" / "Find the event with Müller" | `search-events --query "Müller"` (subject search across ALL calendars) OR `events --past N --all-calendars` (date range) |
| "Show me the details of event X" | `event-detail EVENT_ID` (full body, attendees, recurrence, categories) |
| "What was the last change?" / "Who created X?" | `audit-log --tail 5` (forensics from the write-op log) |
| "Add X on DD.MM. at HH:MM" | `create-event` (after confirmation, automatic idempotency) |
| "Move X to tomorrow" / "Rename X" | `update-event` (after confirmation) |
| "Delete event X" | `delete-event` (after confirmation, with event ID from last `events` output) |
| "Which tasks are open?" | `tasks` |
| "Put X on the list" / "Remind me about X" | `create-task` (after confirmation) |
| "Change task X" / "Move the due date of X" | `update-task` (after confirmation) |
| "I finished X" | `complete-task` (after confirmation) |
| "Delete task X" | `delete-task` (after confirmation) |
| "Which to-do lists do I have?" | `lists` |
| "Make a new list X" / "Rename list X to Y" / "Delete list X" | `create-list` / `rename-list` / `delete-list` (after confirmation) |
| "Which steps does task X have?" | `steps` |
| "Add step X to task Y" / "Step X done" / "Delete step X" | `add-step` / `complete-step` / `delete-step` (after confirmation) |
| "Is Outlook still connected?" / "How long is the token valid?" | `status` |

## Microsoft To Do: lists vs. app groups

In the Microsoft To Do app sidebar you see:
- **Lists** (each colored entry, e.g. `🏠 Household`, `📥 Inbox`)
- Sometimes several lists collapsed under a "Group" header

**Worth knowing:**
- What the Graph API calls a `todoTaskList` = the list level (each of those entries)
- The "Group" headers that visually bundle several lists are **client-side app UI** and not accessible in Graph v1.0 or Beta. Microsoft likely stores them per-app-install locally
- The old `outlookTaskGroup`/`outlookTaskFolder` API with a real folder hierarchy was deprecated on 2022-08-20 and returns no data anymore

**Consequence:** Users often call the **lists** "folders" — and those we can fully manage. What we CANNOT do: programmatically build a "Group" around several lists (that's UI state the user sets via drag-and-drop in the app).

**List CRUD:** `create-list`, `rename-list`, `delete-list`. Built-in lists (`wellknownListName=defaultList`, `Flagged Emails`) are protected; override with `--allow-default`, though Graph would reject it anyway.

## Subcommands

### `calendars` — list all calendars

```bash
outlook.py calendars
```

Returns `id`, `name`, `color`, `is_default`, `can_edit`, `owner` per calendar. Use the IDs/names to set `default_calendar` in `config.json` or to target a specific calendar via `--calendar`.

### `events` — read calendar

```bash
outlook.py events                       # next 7 days, default calendar only
outlook.py events --all-calendars       # next 7 days, across ALL calendars
outlook.py events --days 1              # today only, default calendar
outlook.py events --days 14 --all-calendars
outlook.py events --past 14 --days 0    # the last 14 days (lookback)
outlook.py events --past 30 --days 30 --all-calendars   # +/- 30 days
outlook.py events --calendar "Work"     # a specific calendar by name
outlook.py events --calendar all        # synonym for --all-calendars
```

**No `--start`/`--end` flag** — `events` only knows `--days` (forward) and `--past` (backward). For ranges use e.g. `--past 7 --days 0`.

**Note:** an early bug made `--all-calendars` user-default-only (it ignored shared calendars). It now iterates `/me/calendars` and calls `/me/calendars/{id}/calendarView` per calendar — same pattern as `search-events`. Output carries a `calendar` field per event for provenance.

**Important:** the default is ONLY the `default_calendar` from `config.json`. An event sitting in another calendar would NOT show up with a bare `events` call. For day/week overviews, **always pass `--all-calendars`**.

Output: `{ "calendar": "...", "past_days": N, "future_days": N, "count": M, "events": [...] }` — each event has `id, subject, start, end, all_day, location, organizer, body_preview, web_link, type, series_master_id`.

### `create-event` — create an event

Required: `--subject`, `--start`, `--end`. ISO-8601 local, e.g. `2026-05-20T18:00:00`.

**⚠️ Pitfall: the flag is `--subject`, NOT `--summary`** — `--summary` is the Google-Calendar name. Here it's always `--subject`.

```bash
# Dry run first — show the draft to the user:
outlook.py create-event --subject "Dentist" \
  --start "2026-05-20T10:00:00" --end "2026-05-20T11:00:00" \
  --location "Main St 12" --dry-run

# After confirmation (drop --dry-run):
outlook.py create-event --subject "Dentist" \
  --start "2026-05-20T10:00:00" --end "2026-05-20T11:00:00" \
  --location "Main St 12"
```

Default calendar = `default_calendar` from `config.json`. Target another via `--calendar "Name"`.

**Idempotency:** `create-event` automatically sets a content-based `transactionId` (SHA-256 of subject + start + end + calendar id, truncated to 40 chars). Microsoft Graph dedupes POSTs with the same transactionId within **6 hours** onto the same event ID. Effect: if the agent retries the call after a network timeout, there's NO duplicate. To intentionally create the same event twice, vary at least one field.

**Validation:**
- `end > start` is checked before the POST — an 18→17 input aborts with a clear error instead of relying on Graph's response
- **Timezone-naive ISO enforced**: `--start`/`--end` must NOT carry a `Z` suffix or `+02:00`/`-05:00` offset. The TZ is set separately via the `Prefer: outlook.timezone="..."` header. The script rejects TZ-aware input explicitly — mixing offset + header TZ is ambiguous (off-by-one in winter if the caller hardcodes a summer offset).

**Idempotency edge on delete+recreate** (verified): after `delete-event` you can immediately create a fresh event with the SAME content-based transactionId — Graph holds no 6h tombstone. The 6h dedup window only guards against retry-after-timeout, not against deliberate user intent.

### `lists` — show To Do lists

```bash
outlook.py lists
```

Returns all lists with `id`, `name`, `wellknown` (= "defaultList" for the built-in "Tasks" list).

### `tasks` — read tasks

```bash
outlook.py tasks                        # open tasks in the default list
outlook.py tasks --all                  # incl. completed
outlook.py tasks --list "Shopping"
```

Output: `{ "list": "...", "count": N, "tasks": [...] }` — each task has `id, title, status, importance, due, body`.

### `create-task` — create a task

```bash
outlook.py create-task --title "Swap tyres" \
  --due "2026-05-22T08:00:00" --importance high --dry-run

outlook.py create-task --title "Bananas" --list "Shopping"
```

### `complete-task` — mark task done

```bash
outlook.py complete-task <TASK_ID> --list "Shopping"
```

### `delete-task` — delete a task

```bash
outlook.py delete-task <TASK_ID> --list "Shopping" --dry-run
outlook.py delete-task <TASK_ID> --list "Shopping"
```

### `delete-event` — delete an event

```bash
outlook.py delete-event <EVENT_ID> --dry-run
outlook.py delete-event <EVENT_ID>
```

Get event IDs from `events` output (`id` field). The dry run shows subject + start/end so the user can confirm it's the right event.

**For recurring events (birthdays etc.):** if the event is a series master, the dry run emits a `WARNING` that the whole series will be deleted — show that warning before confirming. To delete just ONE occurrence, the `events` output has the single occurrence with `type: occurrence` + `series_master_id` — then DELETE the occurrence ID; the rest of the series stays.

### `update-event` — change an event (PATCH)

```bash
# Move:
outlook.py update-event <EVENT_ID> --start "2026-05-21T11:00:00" --end "2026-05-21T12:00:00" --dry-run

# Rename:
outlook.py update-event <EVENT_ID> --subject "Dentist Dr. Smith"

# Several fields at once:
outlook.py update-event <EVENT_ID> --location "Main St 12" --body "New address"
```

At least one field is required. `--dry-run` shows the patch body, `--calendar` if the event lives in a non-default calendar.

**Careful with series masters:** PATCH on a seriesMaster changes ALL future occurrences. For a single-occurrence change, pull the occurrence from `events` first and patch its ID — Graph then creates an "exception" automatically.

### `update-task` — change a task (PATCH)

```bash
# Change title + move due date:
outlook.py update-task <TASK_ID> --title "new title" --due "2026-06-15T10:00:00"

# Clear the due date (empty string):
outlook.py update-task <TASK_ID> --due ""

# Set status (notStarted | inProgress | completed | waitingOnOthers | deferred):
outlook.py update-task <TASK_ID> --status inProgress

# Raise importance:
outlook.py update-task <TASK_ID> --importance high
```

At least one field is required. `--list` if not in the default list.

### Recurring events — `--recurrence` (birthdays, weekly, monthly, etc.)

Graph accepts four pattern types: `daily`, `weekly`, `monthly` (= absoluteMonthly, same day of month), `yearly` (= absoluteYearly, same day & month).

```bash
# Classic birthday (yearly, endless):
outlook.py create-event --subject "🎂 Mum" \
  --start "2026-04-15T00:00:00" --end "2026-04-16T00:00:00" \
  --all-day --recurrence yearly

# Weekly on Mon + Thu, with an end date:
outlook.py create-event --subject "Gym" \
  --start "2026-05-18T18:00:00" --end "2026-05-18T19:30:00" \
  --recurrence weekly --recurrence-days "mon,thu" \
  --recurrence-until "2026-12-31"

# Every month on the 1st, for 12 months:
outlook.py create-event --subject "Check rent" \
  --start "2026-06-01T09:00:00" --end "2026-06-01T09:15:00" \
  --recurrence monthly --recurrence-count 12

# Every 3 days:
outlook.py create-event --subject "Water plants" \
  --start "2026-05-20T08:00:00" --end "2026-05-20T08:05:00" \
  --recurrence daily --recurrence-interval 3
```

**Flags:**
- `--recurrence {daily,weekly,monthly,yearly}` — pattern type. Required to turn recurrence on
- `--recurrence-interval N` — every N units (default 1). `yearly --recurrence-interval 4` = every 4 years
- `--recurrence-days mon,tue,...` — weekly only. If omitted, the weekday is derived from `--start`. Accepts English and German abbreviations (`mo,di,mi,do,fr,sa,so`)
- `--recurrence-until YYYY-MM-DD` OR `--recurrence-count N` — end of the series. Without either: runs endless (`noEnd`)

**For birthdays:** set `--all-day` together with `--recurrence yearly`. Start = `YYYY-MM-DDT00:00:00`, end = next day; the script forces midnight when `--all-day` is set.

The same pattern works for tasks:
```bash
outlook.py create-task --title "Bins out" \
  --due "2026-05-21T19:00:00" \
  --recurrence weekly --recurrence-days "thu"
```

### List CRUD

```bash
outlook.py create-list --name "🏠 Household – Shopping" --dry-run
outlook.py create-list --name "🏠 Household – Shopping"

outlook.py rename-list --list "Shopping list" --new-name "🛒 Shopping"
outlook.py delete-list --list "Old list" --dry-run    # shows open tasks
outlook.py delete-list --list "Old list"
```

`delete-list` blocks built-in lists (defaultList, flaggedEmails) via the `wellknownListName` check.

### Subtasks / steps (`steps`, `add-step`, `complete-step`, `delete-step`)

In Microsoft To Do these are "steps", in Graph `checklistItems`.

```bash
outlook.py steps <TASK_ID> --list "Shopping"
outlook.py add-step <TASK_ID> --title "Scan receipt" --list "Shopping"
outlook.py complete-step <TASK_ID> <STEP_ID> --list "Shopping"
outlook.py delete-step <TASK_ID> <STEP_ID> --list "Shopping" --dry-run
outlook.py delete-step <TASK_ID> <STEP_ID> --list "Shopping"
```

### `search-events` — subject keyword search

```bash
outlook.py search-events --query "Müller"           # default limit 25
outlook.py search-events --query "Dentist" --limit 5
```

Searches ALL calendars (iterates `/me/calendars/{id}/events?$filter=contains(subject,...)`). Output carries a `calendar` field per hit, sorted by start time (newest first).

**Note:** Graph `$search` is unsupported on `/me/events` (returns 501). We use `$filter contains(subject,...)` as a case-insensitive substring search. Body content is NOT searched (too expensive).

### `event-detail` — full object of one event

```bash
outlook.py event-detail <EVENT_ID>
outlook.py event-detail <EVENT_ID> --calendar "Work"   # if not on the default
```

Returns: `id, subject, type, series_master_id, is_all_day, start, end, time_zone, location, organizer, attendees[], categories[], recurrence, body, body_type, created, last_modified, web_link`.

Use cases:
- Before `update-event`: check whether attendees are present (meeting-update emails)
- Read the full body (the `events` output truncates to 200 chars)
- `created`/`last_modified` for forensics

### `audit-log` — write-op forensics

```bash
outlook.py audit-log                           # default --tail 20
outlook.py audit-log --tail 5
outlook.py audit-log --tail 0                  # all entries
outlook.py audit-log --op delete-event         # delete ops only
outlook.py audit-log --tail 50 --op create-event
```

Reads `audit.log` (append-only JSON-Lines, one per write op). Fields per entry: `ts` (UTC ISO), `op`, plus contextual fields (id, subject, list, task_id, ...).

Use cases:
- "What got created yesterday?" → `audit-log --tail 50 --op create-event`
- "Delete what was just created" → `audit-log --tail 1 --op create-event` → pull the ID → `delete-event`
- On suspicion of an unwanted change: `audit-log --tail 100`

### `status` — health check

```bash
outlook.py status
```

Output:
- `graph_user` / `graph_mail` — who we are according to MS
- `default_calendar` — from `config.json`
- `token_valid_for_s` — seconds until access-token expiry (auto-refresh kicks in at 0)
- `scope` — current token scopes

Useful when an Outlook call fails and it's unclear whether it's auth or another error.

## Rules (IMPORTANT)

**Writes:**
- Before every `create-*`, `delete-*`, `complete-*`, `rename-list`, `add-step`: show a draft and get confirmation. NEVER write blindly.
- Use `--dry-run` when unsure — it returns the request body without side effects.
- After a successful write: briefly show the result (ID, webLink, list name).

**Errors:**
- Graph error: report the stderr body verbatim, do NOT retry-loop or guess.
- `AADSTS70008 refresh token expired`: re-run `setup_auth.py` to re-authenticate.

**Files:**
- Never hand-edit `.credentials/credentials.json` — the script maintains it autonomously on refresh; an atomic rename guards against corruption.
- `audit.log` — append-only JSON-Lines for every write. Format: `{"ts":"ISO-UTC","op":"name","id":...,...}`. Useful forensics for unintended changes.

## Configuration

`.credentials/config.json`:
```json
{
  "client_id": "your-azure-app-client-id",
  "default_calendar": "AAMk... (calendar ID, optional)",
  "default_calendar_name": "Calendar"
}
```

`.credentials/credentials.json` (auto-managed by `setup_auth.py` + refresh):
- `access_token`, `expires_in`, `expires_at` — reset on every refresh
- `refresh_token` — rotated on refresh (if Microsoft issues a new one)
- `scope` — Calendars.ReadWrite{,.Shared}, Tasks.ReadWrite{,.Shared}, User.Read
