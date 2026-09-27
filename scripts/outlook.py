#!/usr/bin/env python3
"""Outlook / Microsoft 365 skill — Calendar + To Do via Microsoft Graph.

Speaks directly to Microsoft Graph for:
  - Calendar read/write (default calendar configurable in config.json)
  - Microsoft To Do tasks (list, create, complete)

Auth: reads/writes ./../.credentials/credentials.json (sibling of scripts/).
Auto-refreshes access_token via login.microsoftonline.com when expired.
Run setup_auth.py once to bootstrap credentials.json from your client_id.

All output is JSON for the agent to parse.  Errors go to stderr + exit 1.

Usage:
  outlook.py events [--days N] [--calendar NAME_OR_ID]
  outlook.py create-event --subject S --start ISO --end ISO
                          [--calendar X] [--body B] [--location L] [--dry-run]
  outlook.py lists
  outlook.py tasks [--list NAME_OR_ID] [--all]
  outlook.py create-task --title T [--list X] [--due ISO] [--body B]
                         [--importance low|normal|high] [--dry-run]
  outlook.py complete-task TASK_ID [--list NAME_OR_ID]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import requests

SKILL_DIR = Path(__file__).resolve().parent.parent
CRED_FILE = SKILL_DIR / ".credentials" / "credentials.json"
CONF_FILE = SKILL_DIR / ".credentials" / "config.json"
AUDIT_LOG = SKILL_DIR / "audit.log"

GRAPH = "https://graph.microsoft.com/v1.0"
# {tenant} = common | consumers | organizations | <tenant-id>, set via
# config.json "tenant" (default "common").  Must match the app's
# signInAudience: "common" needs an app that allows org + personal accounts;
# a personal-accounts-only app needs "consumers".
TOKEN_URL_TMPL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
TZ = os.environ.get("OUTLOOK_TZ", "UTC")

# Graph IDs for calendar items vs. todo items both look like long
# base64-ish blobs.  We treat any string ≥ 60 chars OR starting with the
# known prefixes as an ID; everything else as a display-name lookup.
# Calendars: 'AAMk…'.  Todo lists: 'AQMk…' (most) or 'AAMk…' (built-in).
_ID_PREFIXES = ("AAMk", "AQMk")


def _looks_like_id(s: str) -> bool:
    return s.startswith(_ID_PREFIXES) and len(s) >= 60


# TZ-aware suffixes: 'Z', '+02:00', '-0500'.  Must NOT be present in
# --start/--end because Graph wants a naive dateTime plus a separate
# timeZone field; mixing both makes the result ambiguous (off-by-one
# when the caller hardcodes a summer offset).
_TZ_SUFFIX_RE = re.compile(r"(Z|[+-]\d{2}:?\d{2})$")


def _validate_naive_iso(iso: str, field: str) -> None:
    if iso and _TZ_SUFFIX_RE.search(iso):
        raise SystemExit(
            f"--{field}={iso!r} carries a timezone suffix (Z or +HH:MM). "
            f"Pass a naive ISO 8601 string like '2026-05-20T18:00:00' — "
            f"the timezone is set separately (UTC by default, override via OUTLOOK_TZ)."
        )


def _validate_range(start_iso: str, end_iso: str, kind: str = "event") -> None:
    """end must be strictly after start.  Graph would reject some combos
    silently or with cryptic errors, so do the check up-front.
    Also enforces naive ISO (no Z / no offset) — see _validate_naive_iso."""
    _validate_naive_iso(start_iso, "start")
    _validate_naive_iso(end_iso, "end")
    try:
        s = datetime.fromisoformat(start_iso[:19])
        e = datetime.fromisoformat(end_iso[:19])
    except ValueError as ex:
        raise SystemExit(f"invalid ISO datetime: {ex}")
    if e <= s:
        raise SystemExit(
            f"{kind}: end {end_iso!r} must be strictly after start {start_iso!r}"
        )


def _txid(*parts: str) -> str:
    """Build a deterministic transactionId from event content (subject,
    start, end, calendar) so an accidental retry of the same POST gets
    deduplicated by Graph's 6h window instead of creating a duplicate.

    Max 40 chars per Graph spec; SHA-256 hex truncated to 40 fits."""
    h = hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()
    return h[:40]


def _audit(op: str, **kw) -> None:
    """One JSON line per write-operation, append-only.
    The agent's stdout still sees the normal JSON output;
    this is for forensic 'who modified what when'."""
    try:
        with AUDIT_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(
                {"ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                 "op": op, **kw},
                ensure_ascii=False) + "\n")
    except OSError:
        pass  # never let logging break a real op


# ---------------------------------------------------------------- auth ----

def _load_json(path: Path) -> dict:
    with path.open() as f:
        return json.load(f)


def _save_creds(creds: dict) -> None:
    tmp = CRED_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(creds, indent=2))
    os.chmod(tmp, 0o600)
    tmp.replace(CRED_FILE)


def _refresh(creds: dict, conf: dict) -> dict:
    r = requests.post(
        TOKEN_URL_TMPL.format(tenant=conf.get("tenant", "common")),
        data={
            "grant_type": "refresh_token",
            "refresh_token": creds["refresh_token"],
            "client_id": conf["client_id"],
            "scope": creds.get(
                "scope",
                "offline_access User.Read Calendars.ReadWrite "
                "Calendars.ReadWrite.Shared Tasks.ReadWrite Tasks.ReadWrite.Shared",
            ),
        },
        timeout=20,
    )
    if r.status_code >= 400:
        sys.stderr.write(f"refresh failed {r.status_code}: {r.text[:500]}\n")
        r.raise_for_status()
    new = r.json()
    creds["access_token"] = new["access_token"]
    creds["expires_in"] = new["expires_in"]
    creds["expires_at"] = int(time.time()) + int(new["expires_in"]) - 60
    if "refresh_token" in new:
        creds["refresh_token"] = new["refresh_token"]
    if "scope" in new:
        creds["scope"] = new["scope"]
    _save_creds(creds)
    return creds


def get_token() -> tuple[str, dict]:
    if not CONF_FILE.exists():
        raise SystemExit(
            f"missing {CONF_FILE} — copy .credentials/config.example.json to "
            f"config.json and set your client_id (see README)."
        )
    if not CRED_FILE.exists():
        raise SystemExit(
            f"missing {CRED_FILE} — run `python scripts/setup_auth.py` first "
            f"to authenticate (see README)."
        )
    creds = _load_json(CRED_FILE)
    conf = _load_json(CONF_FILE)
    if int(time.time()) >= int(creds.get("expires_at", 0)):
        creds = _refresh(creds, conf)
    return creds["access_token"], conf


# --------------------------------------------------------------- graph ----

def gapi(method: str, path: str, token: str, **kw):
    headers = kw.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    headers.setdefault("Prefer", f'outlook.timezone="{TZ}"')
    r = requests.request(method, f"{GRAPH}{path}", headers=headers, timeout=30, **kw)
    if r.status_code == 401:
        first_body = r.text[:300]
        try:
            creds = _load_json(CRED_FILE)
            conf = _load_json(CONF_FILE)
            creds = _refresh(creds, conf)
        except Exception as ex:
            sys.stderr.write(
                f"Graph {method} {path} -> 401, AND refresh failed: {ex}\n"
                f"original-401-body: {first_body}\n"
            )
            sys.exit(1)
        headers["Authorization"] = f"Bearer {creds['access_token']}"
        r = requests.request(method, f"{GRAPH}{path}", headers=headers, timeout=30, **kw)
    if r.status_code >= 400:
        sys.stderr.write(f"Graph {method} {path} -> {r.status_code}\n{r.text[:500]}\n")
        r.raise_for_status()
    return r.json() if r.text else {}


_WEEKDAY_MAP = {
    "mon": "monday", "tue": "tuesday", "wed": "wednesday", "thu": "thursday",
    "fri": "friday", "sat": "saturday", "sun": "sunday",
    "mo": "monday", "di": "tuesday", "mi": "wednesday", "do": "thursday",
    "fr": "friday", "sa": "saturday", "so": "sunday",
}


def _normalize_weekdays(raw: str | None, fallback_iso: str | None) -> list[str]:
    """Comma-separated 'mon,thu' → ['monday','thursday'].
    Falls leer und fallback_iso gesetzt: weekday-Name aus dem ISO-Datum ableiten."""
    if raw:
        out = []
        for d in raw.split(","):
            key = d.strip().lower()[:3]
            if key not in _WEEKDAY_MAP:
                raise SystemExit(f"unknown weekday {d!r}.  Use mon|tue|wed|thu|fri|sat|sun")
            out.append(_WEEKDAY_MAP[key])
        return out
    if fallback_iso:
        names = ["monday", "tuesday", "wednesday", "thursday",
                 "friday", "saturday", "sunday"]
        return [names[datetime.fromisoformat(fallback_iso[:10]).weekday()]]
    return []


def _build_recurrence(spec: str | None, start_iso: str, *,
                      interval: int = 1,
                      days_raw: str | None = None,
                      until: str | None = None,
                      count: int | None = None) -> dict | None:
    """Build a patternedRecurrence (works for both event + task).
    start_iso: 'YYYY-MM-DD[Thh:mm:ss]' — first 10 chars used as date anchor."""
    if not spec:
        return None
    spec = spec.lower()
    start_date = start_iso[:10]

    pattern: dict = {"interval": interval}
    if spec == "daily":
        pattern["type"] = "daily"
    elif spec == "weekly":
        pattern["type"] = "weekly"
        pattern["daysOfWeek"] = _normalize_weekdays(days_raw, start_iso)
        pattern["firstDayOfWeek"] = "monday"
    elif spec in ("monthly", "absolutemonthly"):
        pattern["type"] = "absoluteMonthly"
        pattern["dayOfMonth"] = int(start_date[8:10])
    elif spec in ("yearly", "absoluteyearly"):
        pattern["type"] = "absoluteYearly"
        pattern["dayOfMonth"] = int(start_date[8:10])
        pattern["month"] = int(start_date[5:7])
    else:
        raise SystemExit(
            f"unknown recurrence {spec!r}.  Use daily|weekly|monthly|yearly."
        )

    range_: dict = {"type": "noEnd", "startDate": start_date}
    if until and count:
        raise SystemExit("--recurrence-until and --recurrence-count are mutually exclusive")
    if until:
        range_ = {"type": "endDate", "startDate": start_date, "endDate": until[:10]}
    elif count:
        range_ = {"type": "numbered", "startDate": start_date,
                  "numberOfOccurrences": int(count)}
    return {"pattern": pattern, "range": range_}


def _resolve_calendar(arg: str | None, token: str, conf: dict) -> str:
    if not arg:
        cal = conf.get("default_calendar")
        if cal:
            return cal
        # No default configured — fall back to the account's primary calendar
        # so a bare `events`/`create-event` works out of the box.
        return gapi("GET", "/me/calendar?$select=id", token)["id"]
    if _looks_like_id(arg):
        return arg
    cals = gapi("GET", "/me/calendars?$top=100", token).get("value", [])
    for c in cals:
        if c["name"].lower() == arg.lower():
            return c["id"]
    raise SystemExit(
        f"calendar {arg!r} not found.  Known: {[c['name'] for c in cals]}"
    )


def _resolve_list(arg: str | None, token: str) -> tuple[str, str]:
    lists = gapi("GET", "/me/todo/lists?$top=100", token).get("value", [])
    if not lists:
        raise SystemExit("no To Do lists found on this account")
    if not arg:
        for lst in lists:
            if lst.get("wellknownListName") == "defaultList":
                return lst["id"], lst["displayName"]
        return lists[0]["id"], lists[0]["displayName"]
    if _looks_like_id(arg):
        # Look up the real display name so audit-log + output are readable.
        for lst in lists:
            if lst["id"] == arg:
                return arg, lst["displayName"]
        # ID points to a list we don't see (e.g. recently deleted) — let
        # the caller hit Graph and get a clear 404 with the ID echoed.
        return arg, "(unresolved)"
    for lst in lists:
        if lst["displayName"].lower() == arg.lower():
            return lst["id"], lst["displayName"]
    raise SystemExit(
        f"list {arg!r} not found.  Known: {[lst['displayName'] for lst in lists]}"
    )


# ------------------------------------------------------------ commands ----

def _fetch_calendar_view(cal_id: str, cal_name: str, start_z: str, end_z: str,
                         token: str) -> list:
    """Fetch /me/calendars/{cal_id}/calendarView between start_z and end_z.
    Returns a list of formatted event dicts (with the ``calendar`` name attached
    so aggregated multi-calendar output can show provenance)."""
    path = (
        f"/me/calendars/{cal_id}/calendarView"
        f"?startDateTime={start_z}&endDateTime={end_z}"
        f"&$orderby=start/dateTime&$top=100"
        f"&$select=id,subject,start,end,location,bodyPreview,isAllDay,organizer,webLink,type,seriesMasterId"
    )
    data = gapi("GET", path, token)
    out = []
    for e in data.get("value", []):
        out.append(
            {
                "id": e["id"],
                "calendar": cal_name,
                "subject": e["subject"],
                "start": e["start"]["dateTime"][:16].replace("T", " "),
                "end": e["end"]["dateTime"][:16].replace("T", " "),
                "all_day": e["isAllDay"],
                "location": (e.get("location") or {}).get("displayName") or None,
                "organizer": ((e.get("organizer") or {}).get("emailAddress") or {}).get("name"),
                "body_preview": (e.get("bodyPreview") or "")[:200],
                "web_link": e.get("webLink"),
                # type: singleInstance | occurrence | exception | seriesMaster
                "type": e.get("type"),
                "series_master_id": e.get("seriesMasterId"),
            }
        )
    return out


def cmd_events(args):
    token, conf = get_token()
    if args.days < 0 or args.past < 0:
        raise SystemExit("--days and --past must both be >= 0")
    now_local = datetime.now(timezone.utc).astimezone()
    # `--past N` shifts the window back N days; `--days N` is the forward
    # window length.  By default: today through +days_days.
    start = (now_local - timedelta(days=args.past)).replace(
        hour=0, minute=0, second=0, microsecond=0)
    end = start.replace(hour=0, minute=0, second=0, microsecond=0) + \
          timedelta(days=args.past + args.days)
    start_z = start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    end_z = end.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if args.all_calendars or (args.calendar and args.calendar.lower() == "all"):
        # /me/calendarView is user-default-only and silently drops shared
        # calendars — iterate over /me/calendars per-id (same pattern as
        # cmd_search_events) to actually cover ALL owned + shared calendars.
        cals = gapi("GET", "/me/calendars?$top=100&$select=id,name", token).get("value", [])
        out = []
        for c in cals:
            out.extend(_fetch_calendar_view(c["id"], c["name"], start_z, end_z, token))
        # Merge-sort: events come back per-calendar in start order; aggregated
        # view needs a global sort by start time.
        out.sort(key=lambda r: r.get("start") or "")
        calendar_label = "ALL"
    else:
        cal_id = _resolve_calendar(args.calendar, token, conf)
        calendar_label = args.calendar or conf.get("default_calendar_name") or cal_id
        out = _fetch_calendar_view(cal_id, calendar_label, start_z, end_z, token)
    print(json.dumps({"calendar": calendar_label, "past_days": args.past,
                      "future_days": args.days, "count": len(out),
                      "events": out},
                     indent=2, ensure_ascii=False))


def cmd_search_events(args):
    """Subject keyword search via Graph $filter contains().
    Graph does NOT support $search on /me/events (returns 501).  Also,
    /me/events only sees the user-level default calendar — so to cover
    shared calendars as well as the private one, we iterate over
    /me/calendars/{id}/events for each owned calendar."""
    token, _ = get_token()
    # OData string escape: double the apostrophes.
    q = args.query.replace("'", "''")
    fq = quote(q)
    cals = gapi("GET", "/me/calendars?$top=100&$select=id,name", token).get("value", [])
    out: list = []
    for c in cals:
        path = (f"/me/calendars/{c['id']}/events"
                f"?$filter=contains(subject,'{fq}')"
                f"&$top={args.limit}"
                f"&$orderby=start/dateTime desc"
                f"&$select=id,subject,start,end,location,bodyPreview,type,seriesMasterId,webLink")
        data = gapi("GET", path, token)
        for e in data.get("value", []):
            out.append({
                "id": e["id"],
                "calendar": c["name"],
                "subject": e["subject"],
                "start": (e.get("start") or {}).get("dateTime", "")[:16].replace("T", " "),
                "end": (e.get("end") or {}).get("dateTime", "")[:16].replace("T", " "),
                "location": (e.get("location") or {}).get("displayName") or None,
                "body_preview": (e.get("bodyPreview") or "")[:200],
                "type": e.get("type"),
                "series_master_id": e.get("seriesMasterId"),
                "web_link": e.get("webLink"),
            })
    # Sort combined results by start time (descending = newest first).
    out.sort(key=lambda r: r["start"] or "", reverse=True)
    # Trim to overall limit after merge.
    out = out[: args.limit]
    print(json.dumps({"query": args.query, "calendars_searched": len(cals),
                      "count": len(out), "events": out},
                     indent=2, ensure_ascii=False))


def cmd_event_detail(args):
    """Single-event full GET — useful when bodyPreview was truncated or
    you need attendees/categories/recurrence/createdDateTime."""
    token, conf = get_token()
    cal_id = _resolve_calendar(args.calendar, token, conf)
    ev = gapi("GET", f"/me/calendars/{cal_id}/events/{args.event_id}", token)
    # Slim down to the useful fields; raw Graph payload has dozens.
    out = {
        "id": ev["id"],
        "subject": ev.get("subject"),
        "type": ev.get("type"),
        "series_master_id": ev.get("seriesMasterId"),
        "is_all_day": ev.get("isAllDay"),
        "start": (ev.get("start") or {}).get("dateTime"),
        "end": (ev.get("end") or {}).get("dateTime"),
        "time_zone": (ev.get("start") or {}).get("timeZone"),
        "location": (ev.get("location") or {}).get("displayName") or None,
        "organizer": ((ev.get("organizer") or {}).get("emailAddress") or {}).get("name"),
        "attendees": [
            {"name": (a.get("emailAddress") or {}).get("name"),
             "address": (a.get("emailAddress") or {}).get("address"),
             "type": a.get("type"),
             "response": (a.get("status") or {}).get("response")}
            for a in (ev.get("attendees") or [])
        ],
        "categories": ev.get("categories") or [],
        "recurrence": ev.get("recurrence"),
        "body": (ev.get("body") or {}).get("content"),
        "body_type": (ev.get("body") or {}).get("contentType"),
        "created": ev.get("createdDateTime"),
        "last_modified": ev.get("lastModifiedDateTime"),
        "web_link": ev.get("webLink"),
    }
    print(json.dumps(out, indent=2, ensure_ascii=False))


def cmd_audit_log(args):
    """Read recent write-ops from audit.log.  Useful for forensics like
    'what got created yesterday' or 'who deleted event X'."""
    if not AUDIT_LOG.exists():
        print(json.dumps({"entries": [], "note": "no audit log yet"},
                         indent=2, ensure_ascii=False))
        return
    lines = AUDIT_LOG.read_text(encoding="utf-8").splitlines()
    if args.tail and args.tail > 0:
        lines = lines[-args.tail:]
    entries = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            entries.append({"_raw": line, "_parse_error": True})
    if args.op:
        entries = [e for e in entries if e.get("op") == args.op]
    print(json.dumps({"count": len(entries), "entries": entries},
                     indent=2, ensure_ascii=False))


def cmd_calendars(args):
    token, _ = get_token()
    cals = gapi("GET",
                "/me/calendars?$top=100&$select=id,name,owner,isDefaultCalendar,canEdit,color",
                token).get("value", [])
    out = []
    for c in cals:
        out.append({
            "id": c["id"],
            "name": c["name"],
            "color": c.get("color"),
            "is_default": c.get("isDefaultCalendar"),
            "can_edit": c.get("canEdit"),
            "owner": ((c.get("owner") or {}).get("name") or
                      (c.get("owner") or {}).get("address")),
        })
    print(json.dumps({"count": len(out), "calendars": out},
                     indent=2, ensure_ascii=False))


def cmd_update_event(args):
    """PATCH event subject/start/end/location/body without losing the id."""
    token, conf = get_token()
    cal_id = _resolve_calendar(args.calendar, token, conf)
    body: dict = {}
    if args.subject is not None:
        body["subject"] = args.subject
    if args.start is not None:
        body["start"] = {"dateTime": args.start, "timeZone": TZ}
    if args.end is not None:
        body["end"] = {"dateTime": args.end, "timeZone": TZ}
    if args.location is not None:
        body["location"] = {"displayName": args.location}
    if args.body is not None:
        body["body"] = {"contentType": "text", "content": args.body}
    if not body:
        raise SystemExit(
            "update-event needs at least one of: "
            "--subject --start --end --location --body"
        )
    # Either-or naive-ISO check (caller may patch just start or just end)
    if args.start:
        _validate_naive_iso(args.start, "start")
    if args.end:
        _validate_naive_iso(args.end, "end")
    if args.start and args.end:
        _validate_range(args.start, args.end, kind="event")
    # Pre-fetch metadata so dry-run shows current state + flags
    # series-master mutations.  Attendees are listed verbatim — the agent
    # decides whether they're the user's own addresses or really external
    # invitees.  Avoids hard-coding email heuristics.
    pre = gapi("GET", f"/me/calendars/{cal_id}/events/{args.event_id}"
               f"?$select=type,subject,attendees", token)
    warnings: list[str] = []
    if pre.get("type") == "seriesMaster":
        warnings.append(
            "this is a seriesMaster — PATCH propagates to ALL future "
            "occurrences.  To change only one occurrence, fetch its "
            "occurrence-id via `events` and PATCH that instead."
        )
    attendees = [
        {"name": (a.get("emailAddress") or {}).get("name"),
         "address": (a.get("emailAddress") or {}).get("address")}
        for a in (pre.get("attendees") or [])
    ]
    if attendees:
        warnings.append(
            f"this event has {len(attendees)} attendee(s) — if any are NOT "
            f"the user's own addresses, PATCH triggers meeting-update emails "
            f"to them.  Verify before confirming."
        )
    if args.dry_run:
        out = {"dry_run": True, "event_id": args.event_id, "patch": body,
               "current_subject": pre.get("subject"),
               "attendees": attendees}
        if warnings:
            out["WARNINGS"] = warnings
        print(json.dumps(out, indent=2, ensure_ascii=False))
        return
    res = gapi("PATCH", f"/me/calendars/{cal_id}/events/{args.event_id}",
               token, json=body)
    _audit("update-event", id=args.event_id, fields=list(body.keys()),
           series_master=(pre.get("type") == "seriesMaster"),
           attendee_count=len(attendees))
    out = {"id": res["id"], "subject": res.get("subject"),
           "webLink": res.get("webLink"),
           "updated_fields": list(body.keys())}
    if warnings:
        out["WARNINGS"] = warnings
    print(json.dumps(out, indent=2, ensure_ascii=False))


def cmd_update_task(args):
    """PATCH task title/due/body/importance/status."""
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    if args.due:  # empty string = clear, skip validation in that case
        _validate_naive_iso(args.due, "due")
    body: dict = {}
    if args.title is not None:
        body["title"] = args.title
    if args.due is not None:
        # Empty string clears the due date
        body["dueDateTime"] = (
            {"dateTime": args.due, "timeZone": TZ} if args.due else None
        )
    if args.body is not None:
        body["body"] = {"content": args.body, "contentType": "text"}
    if args.importance is not None:
        body["importance"] = args.importance
    if args.status is not None:
        body["status"] = args.status
    if not body:
        raise SystemExit(
            "update-task needs at least one of: "
            "--title --due --body --importance --status"
        )
    if args.dry_run:
        print(json.dumps({"dry_run": True, "list": list_name,
                          "task_id": args.task_id, "patch": body},
                         indent=2, ensure_ascii=False))
        return
    res = gapi("PATCH", f"/me/todo/lists/{list_id}/tasks/{args.task_id}",
               token, json=body)
    _audit("update-task", id=args.task_id, list=list_name,
           fields=list(body.keys()))
    print(json.dumps({"id": res["id"], "list": list_name,
                      "title": res.get("title"),
                      "updated_fields": list(body.keys())},
                     indent=2, ensure_ascii=False))


def cmd_delete_step(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    if args.dry_run:
        # Fetch the step so the preview is meaningful.
        meta = gapi("GET",
                    f"/me/todo/lists/{list_id}/tasks/{args.task_id}"
                    f"/checklistItems/{args.step_id}",
                    token)
        print(json.dumps({"dry_run": True, "would_delete": {
            "id": args.step_id, "task_id": args.task_id, "list": list_name,
            "display_name": meta.get("displayName"),
            "is_checked": meta.get("isChecked"),
        }}, indent=2, ensure_ascii=False))
        return
    gapi("DELETE",
         f"/me/todo/lists/{list_id}/tasks/{args.task_id}"
         f"/checklistItems/{args.step_id}",
         token)
    _audit("delete-step", id=args.step_id, list=list_name,
           task_id=args.task_id)
    print(json.dumps({"deleted_step": args.step_id, "list": list_name,
                      "task_id": args.task_id},
                     indent=2, ensure_ascii=False))


def cmd_create_event(args):
    token, conf = get_token()
    cal_id = _resolve_calendar(args.calendar, token, conf)
    start_dt = args.start[:10] + "T00:00:00" if args.all_day else args.start
    end_dt = args.end[:10] + "T00:00:00" if args.all_day else args.end
    _validate_range(start_dt, end_dt, kind="event")
    body = {
        "subject": args.subject,
        "start": {"dateTime": start_dt, "timeZone": TZ},
        "end": {"dateTime": end_dt, "timeZone": TZ},
        # Idempotency: a client retry within Graph's 6h dedup window
        # produces the SAME id instead of a duplicate event.  Hash is
        # content-based so legitimately different events still get
        # different txids.
        "transactionId": _txid(args.subject, start_dt, end_dt, cal_id),
    }
    if args.all_day:
        body["isAllDay"] = True
    if args.body:
        body["body"] = {"contentType": "text", "content": args.body}
    if args.location:
        body["location"] = {"displayName": args.location}
    rec = _build_recurrence(
        args.recurrence, args.start,
        interval=args.recurrence_interval,
        days_raw=args.recurrence_days,
        until=args.recurrence_until,
        count=args.recurrence_count,
    )
    if rec:
        body["recurrence"] = rec
    if args.dry_run:
        print(json.dumps({"dry_run": True, "calendar": cal_id, "body": body},
                         indent=2, ensure_ascii=False))
        return
    res = gapi("POST", f"/me/calendars/{cal_id}/events", token, json=body)
    _audit("create-event", id=res["id"], subject=res.get("subject"),
           start=body["start"]["dateTime"],
           recurring=bool(res.get("recurrence")))
    print(json.dumps({"id": res["id"], "webLink": res.get("webLink"),
                      "subject": res.get("subject"),
                      "recurring": bool(res.get("recurrence"))},
                     indent=2, ensure_ascii=False))


def cmd_delete_event(args):
    token, conf = get_token()
    cal_id = _resolve_calendar(args.calendar, token, conf)
    # Fetch event metadata for the operator to confirm in the dry-run / log.
    meta = gapi("GET", f"/me/calendars/{cal_id}/events/{args.event_id}"
                f"?$select=subject,start,end,isAllDay,type,seriesMasterId,recurrence",
                token)
    ev_type = meta.get("type")
    preview = {
        "id": args.event_id,
        "subject": meta.get("subject"),
        "start": (meta.get("start") or {}).get("dateTime", "")[:16].replace("T", " "),
        "end": (meta.get("end") or {}).get("dateTime", "")[:16].replace("T", " "),
        "all_day": meta.get("isAllDay"),
        "type": ev_type,
    }
    # Loud warning if this DELETE wipes a whole recurring series.
    if ev_type == "seriesMaster":
        preview["WARNING"] = (
            "this is a seriesMaster — DELETE will remove ALL future "
            "occurrences (the whole recurring series).  To delete just one "
            "occurrence, fetch the occurrence's own id via `events` and "
            "delete that instead."
        )
        rec_pat = ((meta.get("recurrence") or {}).get("pattern") or {})
        preview["recurrence_pattern"] = (
            f"{rec_pat.get('type')}/interval={rec_pat.get('interval')}"
            if rec_pat else None
        )
    if args.dry_run:
        print(json.dumps({"dry_run": True, "would_delete": preview},
                         indent=2, ensure_ascii=False))
        return
    gapi("DELETE", f"/me/calendars/{cal_id}/events/{args.event_id}", token)
    _audit("delete-event", **preview)
    print(json.dumps({"deleted": preview}, indent=2, ensure_ascii=False))


def cmd_lists(args):
    token, _ = get_token()
    lists = gapi("GET", "/me/todo/lists?$top=100", token).get("value", [])
    print(
        json.dumps(
            [
                {"id": lst["id"], "name": lst["displayName"],
                 "wellknown": lst.get("wellknownListName")}
                for lst in lists
            ],
            indent=2,
            ensure_ascii=False,
        )
    )


def cmd_tasks(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    filter_ = "" if args.all else "&$filter=status ne 'completed'"
    data = gapi(
        "GET",
        f"/me/todo/lists/{list_id}/tasks?$top=100&$orderby=createdDateTime desc{filter_}",
        token,
    )
    out = []
    for t in data.get("value", []):
        rec = t.get("recurrence")
        rec_summary = None
        if rec:
            p = rec.get("pattern", {})
            rec_summary = f"{p.get('type')}/interval={p.get('interval')}"
        out.append(
            {
                "id": t["id"],
                "title": t["title"],
                "status": t["status"],
                "importance": t.get("importance"),
                "due": (t.get("dueDateTime") or {}).get("dateTime", "")[:16].replace("T", " ") or None,
                "body": (t.get("body") or {}).get("content", "")[:300],
                "recurrence": rec_summary,
            }
        )
    print(json.dumps({"list": list_name, "count": len(out), "tasks": out},
                     indent=2, ensure_ascii=False))


def cmd_create_task(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    if args.due:
        _validate_naive_iso(args.due, "due")
    body = {"title": args.title}
    if args.due:
        body["dueDateTime"] = {"dateTime": args.due, "timeZone": TZ}
    if args.body:
        body["body"] = {"content": args.body, "contentType": "text"}
    if args.importance:
        body["importance"] = args.importance
    # Tasks need a start anchor for recurrence — use --due, else today.
    rec_anchor = args.due or datetime.now().strftime("%Y-%m-%dT00:00:00")
    rec = _build_recurrence(
        args.recurrence, rec_anchor,
        interval=args.recurrence_interval,
        days_raw=args.recurrence_days,
        until=args.recurrence_until,
        count=args.recurrence_count,
    )
    if rec:
        body["recurrence"] = rec
    if args.dry_run:
        print(json.dumps({"dry_run": True, "list": list_name, "body": body},
                         indent=2, ensure_ascii=False))
        return
    res = gapi("POST", f"/me/todo/lists/{list_id}/tasks", token, json=body)
    _audit("create-task", id=res["id"], list=list_name,
           title=res.get("title"),
           recurring=bool(res.get("recurrence")))
    print(json.dumps({"id": res["id"], "list": list_name,
                      "recurring": bool(res.get("recurrence"))},
                     indent=2, ensure_ascii=False))


def cmd_complete_task(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    gapi(
        "PATCH",
        f"/me/todo/lists/{list_id}/tasks/{args.task_id}",
        token,
        json={"status": "completed"},
    )
    _audit("complete-task", id=args.task_id, list=list_name)
    print(json.dumps({"status": "completed", "id": args.task_id, "list": list_name},
                     indent=2, ensure_ascii=False))


def cmd_delete_task(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    # No $select — todoTask collection/item resource rejects it with 400.
    meta = gapi("GET",
                f"/me/todo/lists/{list_id}/tasks/{args.task_id}",
                token)
    preview = {"id": args.task_id, "list": list_name,
               "title": meta.get("title"), "status": meta.get("status")}
    if args.dry_run:
        print(json.dumps({"dry_run": True, "would_delete": preview},
                         indent=2, ensure_ascii=False))
        return
    gapi("DELETE", f"/me/todo/lists/{list_id}/tasks/{args.task_id}", token)
    _audit("delete-task", **preview)
    print(json.dumps({"deleted": preview}, indent=2, ensure_ascii=False))


# ----------- list (folder) CRUD — Graph v1.0 has no "groups" concept,
# only flat todoTaskList resources.  These three commands cover the
# user-visible side of "managing folders".

def cmd_create_list(args):
    token, _ = get_token()
    body = {"displayName": args.name}
    if args.dry_run:
        print(json.dumps({"dry_run": True, "body": body},
                         indent=2, ensure_ascii=False))
        return
    res = gapi("POST", "/me/todo/lists", token, json=body)
    _audit("create-list", id=res["id"], name=res["displayName"])
    print(json.dumps({"id": res["id"], "name": res["displayName"]},
                     indent=2, ensure_ascii=False))


def cmd_rename_list(args):
    token, _ = get_token()
    list_id, old_name = _resolve_list(args.list, token)
    body = {"displayName": args.new_name}
    if args.dry_run:
        print(json.dumps({"dry_run": True, "id": list_id,
                          "rename": {"from": old_name, "to": args.new_name}},
                         indent=2, ensure_ascii=False))
        return
    res = gapi("PATCH", f"/me/todo/lists/{list_id}", token, json=body)
    _audit("rename-list", id=list_id, old_name=old_name, new_name=res["displayName"])
    print(json.dumps({"id": list_id, "old_name": old_name,
                      "new_name": res["displayName"]},
                     indent=2, ensure_ascii=False))


def cmd_delete_list(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    # Detect Graph's built-in defaultList / flaggedEmails by metadata, not by
    # tenant-specific id prefix.  Graph rejects DELETE on those anyway.
    meta = gapi("GET", f"/me/todo/lists/{list_id}", token)
    well_known = meta.get("wellknownListName") or "none"
    if well_known != "none" and not args.allow_default:
        raise SystemExit(
            f"refusing to delete built-in list {list_name!r} "
            f"(wellknownListName={well_known}).  Pass --allow-default to "
            "force-try anyway (Graph will likely reject)."
        )
    # Show what's about to be lost.  No $select — Graph rejects it on this
    # collection ("RequestBroker--ParseUri" 400).  Tasks list is small enough.
    tasks = gapi("GET", f"/me/todo/lists/{list_id}/tasks?$top=100",
                 token).get("value", [])
    preview = {"id": list_id, "name": list_name,
               "task_count": len(tasks),
               "open_tasks": [t["title"] for t in tasks if t["status"] != "completed"][:20]}
    if args.dry_run:
        print(json.dumps({"dry_run": True, "would_delete": preview},
                         indent=2, ensure_ascii=False))
        return
    gapi("DELETE", f"/me/todo/lists/{list_id}", token)
    _audit("delete-list", **preview)
    print(json.dumps({"deleted": preview}, indent=2, ensure_ascii=False))


# ----------- subtasks (Graph: checklistItems, To Do UI: "Schritte") ----

def cmd_steps(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    data = gapi(
        "GET",
        f"/me/todo/lists/{list_id}/tasks/{args.task_id}/checklistItems?$top=100",
        token,
    )
    out = [
        {"id": s["id"], "display_name": s["displayName"],
         "is_checked": s["isChecked"]}
        for s in data.get("value", [])
    ]
    print(json.dumps({"list": list_name, "task_id": args.task_id,
                      "count": len(out), "steps": out},
                     indent=2, ensure_ascii=False))


def cmd_add_step(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    body = {"displayName": args.title}
    if args.dry_run:
        print(json.dumps({"dry_run": True, "list": list_name,
                          "task_id": args.task_id, "body": body},
                         indent=2, ensure_ascii=False))
        return
    res = gapi(
        "POST",
        f"/me/todo/lists/{list_id}/tasks/{args.task_id}/checklistItems",
        token, json=body,
    )
    _audit("add-step", id=res["id"], list=list_name,
           task_id=args.task_id, title=res["displayName"])
    print(json.dumps({"id": res["id"], "title": res["displayName"],
                      "list": list_name, "task_id": args.task_id},
                     indent=2, ensure_ascii=False))


def cmd_complete_step(args):
    token, _ = get_token()
    list_id, list_name = _resolve_list(args.list, token)
    gapi(
        "PATCH",
        f"/me/todo/lists/{list_id}/tasks/{args.task_id}/checklistItems/{args.step_id}",
        token, json={"isChecked": True},
    )
    _audit("complete-step", id=args.step_id, list=list_name,
           task_id=args.task_id)
    print(json.dumps({"checked": True, "id": args.step_id,
                      "task_id": args.task_id, "list": list_name},
                     indent=2, ensure_ascii=False))


# ------------- status: cheap auth/health check, no destructive ops ----

def cmd_status(args):
    token, conf = get_token()
    me = gapi("GET", "/me?$select=displayName,mail,userPrincipalName", token)
    creds = _load_json(CRED_FILE)
    expires_at = int(creds.get("expires_at", 0))
    now = int(time.time())
    print(json.dumps({
        "graph_user": me.get("displayName"),
        "graph_mail": me.get("mail") or me.get("userPrincipalName"),
        "default_calendar": conf.get("default_calendar_name"),
        "token_valid_for_s": max(0, expires_at - now),
        "scope": creds.get("scope", "").split(),
    }, indent=2, ensure_ascii=False))


def build_parser() -> argparse.ArgumentParser:
    """The CLI's argument parser. Shared with scripts/mcp_server.py, which
    drives the same subcommands (validation, audit log, idempotency) as tools."""
    p = argparse.ArgumentParser(description="Outlook/Microsoft 365 skill")
    sub = p.add_subparsers(dest="cmd", required=True)

    e = sub.add_parser("events", help="list calendar events")
    e.add_argument("--days", type=int, default=7,
                   help='forward window in days (default 7)')
    e.add_argument("--past", type=int, default=0,
                   help='also include N days of past events (default 0)')
    e.add_argument("--calendar",
                   help='calendar ID or display name, or "all" to aggregate')
    e.add_argument("--all-calendars", action="store_true",
                   help='aggregate across all calendars (same as --calendar all)')
    e.set_defaults(func=cmd_events)

    sub.add_parser("calendars",
                   help="list all calendars (id, name, color, default flag)"
                   ).set_defaults(func=cmd_calendars)

    se = sub.add_parser("search-events",
                        help='subject keyword search across all calendars')
    se.add_argument("--query", required=True,
                    help='case-insensitive substring matched against event subject')
    se.add_argument("--limit", type=int, default=25)
    se.set_defaults(func=cmd_search_events)

    ed = sub.add_parser("event-detail",
                        help='single-event full GET (body, attendees, recurrence, ...)')
    ed.add_argument("event_id")
    ed.add_argument("--calendar")
    ed.set_defaults(func=cmd_event_detail)

    al = sub.add_parser("audit-log",
                        help='read recent write-ops from audit.log')
    al.add_argument("--tail", type=int, default=20,
                    help='show last N entries (default 20, 0 = all)')
    al.add_argument("--op",
                    help='filter by op name (e.g. create-event, delete-event)')
    al.set_defaults(func=cmd_audit_log)

    ce = sub.add_parser("create-event", help="create calendar event")
    ce.add_argument("--subject", required=True)
    ce.add_argument("--start", required=True, help='ISO 8601 local, e.g. "2026-05-20T18:00:00"')
    ce.add_argument("--end", required=True)
    ce.add_argument("--calendar")
    ce.add_argument("--body")
    ce.add_argument("--location")
    ce.add_argument("--all-day", action="store_true",
                    help='create as all-day event (forces midnight start/end)')
    ce.add_argument("--recurrence",
                    choices=["daily", "weekly", "monthly", "yearly"],
                    help='make this a recurring event (yearly = birthday-pattern)')
    ce.add_argument("--recurrence-interval", type=int, default=1,
                    help='every N units (default 1)')
    ce.add_argument("--recurrence-days",
                    help='for --recurrence weekly: comma list, e.g. "mon,thu"')
    ce.add_argument("--recurrence-until",
                    help='YYYY-MM-DD — recurrence end date (exclusive of --recurrence-count)')
    ce.add_argument("--recurrence-count", type=int,
                    help='number of occurrences (exclusive of --recurrence-until)')
    ce.add_argument("--dry-run", action="store_true")
    ce.set_defaults(func=cmd_create_event)

    de = sub.add_parser("delete-event", help="delete calendar event")
    de.add_argument("event_id")
    de.add_argument("--calendar")
    de.add_argument("--dry-run", action="store_true")
    de.set_defaults(func=cmd_delete_event)

    ue = sub.add_parser("update-event",
                        help="patch calendar event (move, rename, ...)")
    ue.add_argument("event_id")
    ue.add_argument("--subject")
    ue.add_argument("--start", help='new start, ISO 8601 local')
    ue.add_argument("--end", help='new end, ISO 8601 local')
    ue.add_argument("--location")
    ue.add_argument("--body")
    ue.add_argument("--calendar")
    ue.add_argument("--dry-run", action="store_true")
    ue.set_defaults(func=cmd_update_event)

    sub.add_parser("lists", help="list todo lists").set_defaults(func=cmd_lists)

    t = sub.add_parser("tasks", help="list todo tasks")
    t.add_argument("--list", help="list display name or ID (default: built-in Tasks)")
    t.add_argument("--all", action="store_true", help="include completed")
    t.set_defaults(func=cmd_tasks)

    ct = sub.add_parser("create-task", help="create todo task")
    ct.add_argument("--title", required=True)
    ct.add_argument("--list")
    ct.add_argument("--due", help='ISO 8601 local, e.g. "2026-05-22T08:00:00"')
    ct.add_argument("--body")
    ct.add_argument("--importance", choices=["low", "normal", "high"])
    ct.add_argument("--recurrence",
                    choices=["daily", "weekly", "monthly", "yearly"],
                    help='make this a recurring task')
    ct.add_argument("--recurrence-interval", type=int, default=1)
    ct.add_argument("--recurrence-days",
                    help='for --recurrence weekly: comma list, e.g. "mon,thu"')
    ct.add_argument("--recurrence-until")
    ct.add_argument("--recurrence-count", type=int)
    ct.add_argument("--dry-run", action="store_true")
    ct.set_defaults(func=cmd_create_task)

    cc = sub.add_parser("complete-task", help="mark task completed")
    cc.add_argument("task_id")
    cc.add_argument("--list")
    cc.set_defaults(func=cmd_complete_task)

    dt = sub.add_parser("delete-task", help="delete todo task")
    dt.add_argument("task_id")
    dt.add_argument("--list")
    dt.add_argument("--dry-run", action="store_true")
    dt.set_defaults(func=cmd_delete_task)

    ut = sub.add_parser("update-task", help="patch todo task")
    ut.add_argument("task_id")
    ut.add_argument("--list")
    ut.add_argument("--title")
    ut.add_argument("--due", help='ISO 8601 local, or empty string "" to clear')
    ut.add_argument("--body")
    ut.add_argument("--importance", choices=["low", "normal", "high"])
    ut.add_argument("--status",
                    choices=["notStarted", "inProgress", "completed",
                             "waitingOnOthers", "deferred"])
    ut.add_argument("--dry-run", action="store_true")
    ut.set_defaults(func=cmd_update_task)

    # ---- list (folder-like) management ----
    cl = sub.add_parser("create-list", help="create new todo list")
    cl.add_argument("--name", required=True)
    cl.add_argument("--dry-run", action="store_true")
    cl.set_defaults(func=cmd_create_list)

    rl = sub.add_parser("rename-list", help="rename existing todo list")
    rl.add_argument("--list", required=True, help="current list name or ID")
    rl.add_argument("--new-name", required=True)
    rl.add_argument("--dry-run", action="store_true")
    rl.set_defaults(func=cmd_rename_list)

    dl = sub.add_parser("delete-list", help="delete todo list (incl. its tasks)")
    dl.add_argument("--list", required=True, help="list name or ID")
    dl.add_argument("--allow-default", action="store_true",
                    help='allow deleting the built-in default list (Graph will reject anyway)')
    dl.add_argument("--dry-run", action="store_true")
    dl.set_defaults(func=cmd_delete_list)

    # ---- subtasks / "Schritte" / checklistItems ----
    st = sub.add_parser("steps", help="list checklist items (steps) for a task")
    st.add_argument("task_id")
    st.add_argument("--list")
    st.set_defaults(func=cmd_steps)

    ast_ = sub.add_parser("add-step", help="add checklist item to task")
    ast_.add_argument("task_id")
    ast_.add_argument("--title", required=True)
    ast_.add_argument("--list")
    ast_.add_argument("--dry-run", action="store_true")
    ast_.set_defaults(func=cmd_add_step)

    cs = sub.add_parser("complete-step", help="mark checklist item as done")
    cs.add_argument("task_id")
    cs.add_argument("step_id")
    cs.add_argument("--list")
    cs.set_defaults(func=cmd_complete_step)

    ds = sub.add_parser("delete-step", help="delete a checklist item")
    ds.add_argument("task_id")
    ds.add_argument("step_id")
    ds.add_argument("--list")
    ds.add_argument("--dry-run", action="store_true")
    ds.set_defaults(func=cmd_delete_step)

    # ---- diagnostics ----
    sub.add_parser("status", help="show graph identity + token freshness"
                   ).set_defaults(func=cmd_status)
    return p


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except requests.HTTPError:
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as ex:
        sys.stderr.write(f"unexpected: {type(ex).__name__}: {ex}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
