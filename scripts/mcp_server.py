#!/usr/bin/env python3
"""Microsoft 365 Calendar + To Do as an MCP server (stdio).

A thin adapter over the same subcommands as scripts/outlook.py: every tool
builds the CLI arguments and runs the existing command, so validation, the
write audit log, idempotent writes (transactionId) and the token refresh are
exactly the ones the CLI already has - no second implementation to drift.

Safe by default:
  - read-only unless M365_MCP_WRITE=1 (create/update/complete tools appear)
  - deleting needs M365_MCP_DELETE=1 on top (delete tools appear)
  - tools that are not enabled are not registered at all, so a model never
    sees them; every tool carries MCP annotations (readOnlyHint,
    destructiveHint, idempotentHint)

Auth and config are the CLI's: .credentials/config.json + credentials.json
next to scripts/ (run scripts/setup_auth.py once).

Run:  python3 scripts/mcp_server.py            (stdio transport)
Claude Code:
      claude mcp add m365 -- python3 /path/to/scripts/mcp_server.py
      claude mcp add m365 -e M365_MCP_WRITE=1 -- python3 /path/to/scripts/mcp_server.py
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import os
import sys
import threading
from pathlib import Path
from typing import Annotated, Any, Literal

import requests
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

_SCRIPT = Path(__file__).resolve().parent / "outlook.py"
_spec = importlib.util.spec_from_file_location("outlook", _SCRIPT)
outlook = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(outlook)

# The commands print their JSON to stdout - which, over stdio, IS the MCP
# channel. Every call therefore captures stdout/stderr, and one lock keeps
# two tool calls from sharing the (process-wide) redirect.
_LOCK = threading.Lock()

READ = ToolAnnotations(read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=True)
WRITE = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=True, open_world_hint=True)
DELETE = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=True, open_world_hint=True)


def _flag(env: str) -> bool:
    return os.environ.get(env, "").strip().lower() in ("1", "true", "yes", "on")


def run_cli(argv: list[str]) -> dict[str, Any]:
    """Run one outlook.py subcommand and return its JSON; every failure becomes
    a ToolError with the command's own message (the model can read and fix it)."""
    out, err = io.StringIO(), io.StringIO()
    with _LOCK:
        try:
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                args = outlook.build_parser().parse_args(argv)
                args.func(args)
        except SystemExit as ex:
            # a string code is the command's own validation message; an int means
            # argparse or a Graph failure already explained itself on stderr
            msg = ex.code if isinstance(ex.code, str) else err.getvalue().strip()
            raise ToolError(msg or "command failed") from None
        except requests.HTTPError as ex:
            raise ToolError(err.getvalue().strip() or str(ex)) from None
        except requests.RequestException as ex:
            raise ToolError(f"Microsoft Graph not reachable: {type(ex).__name__}") from None
    text = out.getvalue().strip()
    if not text:
        return {"ok": True}
    try:
        data = json.loads(text)
    except ValueError:
        return {"output": text}
    return data if isinstance(data, dict) else {"result": data}


def _opt(argv: list[str], flag: str, value: Any) -> None:
    """Append `flag value` when a value was given (False/None/"" mean: not given)."""
    if value is None or value is False:
        return
    if value is True:
        argv.append(flag)
    else:
        argv += [flag, str(value)]


Recurrence = Literal["daily", "weekly", "monthly", "yearly"]
Importance = Literal["low", "normal", "high"]
# Parameter descriptions end up in the tool's input schema, where the model reads them.
LocalTime = Annotated[str, Field(description='ISO 8601 local time without offset, e.g. "2026-05-20T18:00:00"')]
OptLocalTime = Annotated[str | None, Field(description='ISO 8601 local time without offset, e.g. "2026-05-20T18:00:00"')]
Calendar = Annotated[str | None, Field(description="calendar display name or ID; default: the configured or primary calendar")]
TaskList = Annotated[str | None, Field(description="To Do list display name or ID; default: the built-in Tasks list")]
DryRun = Annotated[bool, Field(description="validate and show what would be sent, without writing")]
Until = Annotated[str | None, Field(description="YYYY-MM-DD end of the recurrence (not together with recurrence_count)")]
Weekdays = Annotated[str | None, Field(description='for weekly recurrence: comma list, e.g. "mon,thu"')]


def create_server(*, allow_write: bool | None = None, allow_delete: bool | None = None) -> MCPServer:
    """Build the server. Without arguments the gates come from the environment."""
    allow_write = _flag("M365_MCP_WRITE") if allow_write is None else allow_write
    allow_delete = _flag("M365_MCP_DELETE") if allow_delete is None else allow_delete
    mode = "read-write-delete" if allow_write and allow_delete else "read-write" if allow_write else "read-only"
    mcp = MCPServer(
        name="m365",
        title="Microsoft 365 Calendar + To Do",
        version="1.0.0",
        instructions=(
            "Microsoft 365 calendar and Microsoft To Do via Microsoft Graph. "
            f"Mode: {mode}. Times are ISO 8601 local times without offset (e.g. 2026-05-20T18:00:00); "
            "calendars and lists accept a display name or an ID. Prefer dry_run=True first for writes."
        ),
    )

    # ---------------------------------------------------------------- read

    @mcp.tool(annotations=READ)
    def status() -> dict:
        """Signed-in Graph identity and token freshness (diagnostics)."""
        return run_cli(["status"])

    @mcp.tool(annotations=READ)
    def list_calendars() -> dict:
        """All calendars (id, name, color, default flag), owned and shared."""
        return run_cli(["calendars"])

    @mcp.tool(annotations=READ)
    def list_events(days: int = 7, past: int = 0, calendar: Calendar = None, all_calendars: bool = False) -> dict:
        """Events from today for `days` days (plus `past` days back). `calendar` is a
        name or ID; all_calendars=True aggregates owned and shared calendars."""
        argv = ["events", "--days", str(days), "--past", str(past)]
        _opt(argv, "--calendar", calendar)
        _opt(argv, "--all-calendars", all_calendars)
        return run_cli(argv)

    @mcp.tool(annotations=READ)
    def search_events(query: str, limit: int = 25) -> dict:
        """Case-insensitive subject search across all calendars."""
        return run_cli(["search-events", "--query", query, "--limit", str(limit)])

    @mcp.tool(annotations=READ)
    def get_event(event_id: str, calendar: Calendar = None) -> dict:
        """One event in full (body, attendees, recurrence, online meeting)."""
        argv = ["event-detail", event_id]
        _opt(argv, "--calendar", calendar)
        return run_cli(argv)

    @mcp.tool(annotations=READ)
    def list_task_lists() -> dict:
        """All Microsoft To Do lists."""
        return run_cli(["lists"])

    @mcp.tool(annotations=READ)
    def list_tasks(task_list: TaskList = None, include_completed: bool = False) -> dict:
        """Tasks of one To Do list (name or ID; default: the built-in Tasks list)."""
        argv = ["tasks"]
        _opt(argv, "--list", task_list)
        _opt(argv, "--all", include_completed)
        return run_cli(argv)

    @mcp.tool(annotations=READ)
    def list_steps(task_id: str, task_list: TaskList = None) -> dict:
        """Checklist items ("steps") of one task."""
        argv = ["steps", task_id]
        _opt(argv, "--list", task_list)
        return run_cli(argv)

    @mcp.tool(annotations=READ)
    def audit_log(tail: int = 20, op: str | None = None) -> dict:
        """Recent write operations from the local audit log (who changed what, when)."""
        argv = ["audit-log", "--tail", str(tail)]
        _opt(argv, "--op", op)
        return run_cli(argv)

    # --------------------------------------------------------------- write

    if allow_write:
        @mcp.tool(annotations=WRITE)
        def create_event(subject: str, start: LocalTime, end: LocalTime, calendar: Calendar = None,
                         body: str | None = None, location: str | None = None, all_day: bool = False,
                         recurrence: Recurrence | None = None, recurrence_interval: int = 1,
                         recurrence_days: Weekdays = None, recurrence_until: Until = None,
                         recurrence_count: int | None = None, dry_run: DryRun = False) -> dict:
            """Create a calendar event (idempotent: a retry within 6 h is deduplicated by Graph).
            recurrence_days for weekly, e.g. "mon,thu"; until (YYYY-MM-DD) and count exclude each other."""
            argv = ["create-event", "--subject", subject, "--start", start, "--end", end,
                    "--recurrence-interval", str(recurrence_interval)]
            for flag, v in (("--calendar", calendar), ("--body", body), ("--location", location),
                            ("--all-day", all_day), ("--recurrence", recurrence),
                            ("--recurrence-days", recurrence_days), ("--recurrence-until", recurrence_until),
                            ("--recurrence-count", recurrence_count), ("--dry-run", dry_run)):
                _opt(argv, flag, v)
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def update_event(event_id: str, subject: str | None = None, start: OptLocalTime = None,
                         end: OptLocalTime = None, location: str | None = None, body: str | None = None,
                         calendar: Calendar = None, dry_run: DryRun = False) -> dict:
            """Change an event (move, rename, ...). Only the given fields change."""
            argv = ["update-event", event_id]
            for flag, v in (("--subject", subject), ("--start", start), ("--end", end),
                            ("--location", location), ("--body", body), ("--calendar", calendar),
                            ("--dry-run", dry_run)):
                _opt(argv, flag, v)
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def create_task(title: str, task_list: TaskList = None, due: OptLocalTime = None, body: str | None = None,
                        importance: Importance | None = None, recurrence: Recurrence | None = None,
                        recurrence_interval: int = 1, recurrence_days: Weekdays = None,
                        recurrence_until: Until = None, recurrence_count: int | None = None,
                        dry_run: DryRun = False) -> dict:
            """Create a To Do task (due as ISO 8601 local time)."""
            argv = ["create-task", "--title", title, "--recurrence-interval", str(recurrence_interval)]
            for flag, v in (("--list", task_list), ("--due", due), ("--body", body), ("--importance", importance),
                            ("--recurrence", recurrence), ("--recurrence-days", recurrence_days),
                            ("--recurrence-until", recurrence_until), ("--recurrence-count", recurrence_count),
                            ("--dry-run", dry_run)):
                _opt(argv, flag, v)
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def update_task(task_id: str, task_list: TaskList = None, title: str | None = None, due: OptLocalTime = None,
                        body: str | None = None, importance: Importance | None = None,
                        status: Literal["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"] | None = None,
                        dry_run: DryRun = False) -> dict:
            """Change a task. Only the given fields change; due="" clears the due date."""
            argv = ["update-task", task_id]
            for flag, v in (("--list", task_list), ("--title", title), ("--body", body),
                            ("--importance", importance), ("--status", status), ("--dry-run", dry_run)):
                _opt(argv, flag, v)
            if due is not None:
                argv += ["--due", due]  # "" is a real value here: clear the due date
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def complete_task(task_id: str, task_list: TaskList = None) -> dict:
            """Mark a task as completed."""
            argv = ["complete-task", task_id]
            _opt(argv, "--list", task_list)
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def add_step(task_id: str, title: str, task_list: TaskList = None, dry_run: DryRun = False) -> dict:
            """Add a checklist item ("step") to a task."""
            argv = ["add-step", task_id, "--title", title]
            _opt(argv, "--list", task_list)
            _opt(argv, "--dry-run", dry_run)
            return run_cli(argv)

        @mcp.tool(annotations=WRITE)
        def complete_step(task_id: str, step_id: str, task_list: TaskList = None) -> dict:
            """Mark a checklist item as done."""
            argv = ["complete-step", task_id, step_id]
            _opt(argv, "--list", task_list)
            return run_cli(argv)

    # -------------------------------------------------------------- delete

    if allow_write and allow_delete:
        @mcp.tool(annotations=DELETE)
        def delete_event(event_id: str, calendar: Calendar = None, dry_run: DryRun = False) -> dict:
            """Delete a calendar event. For a series master this removes the whole series."""
            argv = ["delete-event", event_id]
            _opt(argv, "--calendar", calendar)
            _opt(argv, "--dry-run", dry_run)
            return run_cli(argv)

        @mcp.tool(annotations=DELETE)
        def delete_task(task_id: str, task_list: TaskList = None, dry_run: DryRun = False) -> dict:
            """Delete a To Do task."""
            argv = ["delete-task", task_id]
            _opt(argv, "--list", task_list)
            _opt(argv, "--dry-run", dry_run)
            return run_cli(argv)

    return mcp


def main() -> None:
    if _flag("M365_MCP_DELETE") and not _flag("M365_MCP_WRITE"):
        print("m365 MCP: M365_MCP_DELETE needs M365_MCP_WRITE=1 as well; starting read-only.", file=sys.stderr)
    create_server().run("stdio")


if __name__ == "__main__":
    main()
