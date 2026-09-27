"""MCP server tests: a real MCP client talks to the server in-process; Microsoft
Graph is replaced by a small fake, so no network and no credentials are needed.

    .venv/bin/python -m pytest tests/test_mcp_server.py -q
"""
import asyncio
import importlib.util
import json
import sys
from pathlib import Path

import pytest
import requests
from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("mcp_server", ROOT / "scripts" / "mcp_server.py")
srv = importlib.util.module_from_spec(spec)
spec.loader.exec_module(srv)
o = srv.outlook

READ_TOOLS = {"status", "list_calendars", "list_events", "search_events", "get_event",
              "list_task_lists", "list_tasks", "list_steps", "audit_log"}
WRITE_TOOLS = {"create_event", "update_event", "create_task", "update_task",
               "complete_task", "add_step", "complete_step"}
DELETE_TOOLS = {"delete_event", "delete_task"}


class FakeGraph:
    """Answers the Graph paths the commands use; records every call."""

    def __init__(self):
        self.calls = []
        self.fail = None  # (status, text) → the next call fails like Graph does

    def __call__(self, method, path, token, **kw):
        self.calls.append((method, path, kw.get("json")))
        if self.fail:
            status, text = self.fail
            sys.stderr.write(f"Graph {method} {path} -> {status}\n{text}\n")
            raise requests.HTTPError(f"{status}")
        if path.startswith("/me/calendar?"):
            return {"id": "AAMkPRIMARY"}
        if path.startswith("/me/calendars?"):
            return {"value": [{"id": "AAMkPRIMARY", "name": "Kalender", "isDefaultCalendar": True, "canEdit": True,
                               "color": "auto", "owner": {"name": "Test"}}]}
        if "/calendarView" in path:
            return {"value": [{"id": "EV1", "subject": "Zahnarzt", "start": {"dateTime": "2026-10-01T10:00:00"},
                               "end": {"dateTime": "2026-10-01T11:00:00"}, "isAllDay": False, "type": "singleInstance"}]}
        if method == "POST" and path.endswith("/events"):
            return {"id": "EVNEW", "subject": kw["json"]["subject"], "webLink": "https://outlook.example/e"}
        if path.startswith("/me/todo/lists?"):
            return {"value": [{"id": "AQMkDEFAULT", "displayName": "Aufgaben", "wellknownListName": "defaultList"}]}
        if method == "POST" and path.endswith("/tasks"):
            return {"id": "TNEW", "title": kw["json"]["title"], "status": "notStarted"}
        if method == "PATCH" and "/tasks/" in path:
            return {"id": path.rsplit("/", 1)[-1], "status": "completed"}
        if "/tasks?" in path:
            return {"value": [{"id": "T1", "title": "Bananen", "status": "notStarted", "importance": "normal"}]}
        raise AssertionError(f"unexpected Graph call {method} {path}")


@pytest.fixture
def graph(monkeypatch, tmp_path):
    fake = FakeGraph()
    monkeypatch.setattr(o, "gapi", fake)
    monkeypatch.setattr(o, "get_token", lambda: ("token", {}))
    monkeypatch.setattr(o, "AUDIT_LOG", tmp_path / "audit.log")
    return fake


def run(coro):
    return asyncio.run(coro)


async def _tools(server):
    async with Client(server) as c:
        return {t.name: t for t in (await c.list_tools()).tools}


async def _call(server, name, args=None):
    async with Client(server) as c:
        return await c.call_tool(name, args or {})


def _text(res):
    return "".join(getattr(x, "text", "") for x in res.content)


def _data(res):
    return res.structured_content if res.structured_content else json.loads(_text(res))


# --------------------------------------------------------------- the gates

def test_read_only_by_default(monkeypatch):
    monkeypatch.delenv("M365_MCP_WRITE", raising=False)
    monkeypatch.delenv("M365_MCP_DELETE", raising=False)
    tools = run(_tools(srv.create_server()))
    assert set(tools) == READ_TOOLS
    assert all(t.annotations.read_only_hint for t in tools.values())


def test_write_gate_adds_write_tools_but_not_delete():
    tools = run(_tools(srv.create_server(allow_write=True, allow_delete=False)))
    assert set(tools) == READ_TOOLS | WRITE_TOOLS
    assert not any(tools[n].annotations.destructive_hint for n in WRITE_TOOLS)


def test_delete_needs_both_gates():
    assert set(run(_tools(srv.create_server(allow_write=False, allow_delete=True)))) == READ_TOOLS
    tools = run(_tools(srv.create_server(allow_write=True, allow_delete=True)))
    assert set(tools) == READ_TOOLS | WRITE_TOOLS | DELETE_TOOLS
    assert all(tools[n].annotations.destructive_hint for n in DELETE_TOOLS)


def test_env_flags_open_the_gates(monkeypatch):
    monkeypatch.setenv("M365_MCP_WRITE", "1")
    monkeypatch.setenv("M365_MCP_DELETE", "true")
    assert DELETE_TOOLS <= set(run(_tools(srv.create_server())))


# ------------------------------------------------------------------ reading

def test_list_events_uses_primary_calendar(graph):
    data = _data(run(_call(srv.create_server(), "list_events", {"days": 3})))
    assert data["count"] == 1 and data["events"][0]["subject"] == "Zahnarzt"
    assert any("/me/calendars/AAMkPRIMARY/calendarView" in p for _, p, _ in graph.calls)


def test_list_tasks_resolves_default_list(graph):
    data = _data(run(_call(srv.create_server(), "list_tasks")))
    assert data["list"] == "Aufgaben" and data["tasks"][0]["title"] == "Bananen"
    assert any("AQMkDEFAULT/tasks" in p and "status ne 'completed'" in p for _, p, _ in graph.calls)


def test_nothing_reaches_stdout(graph, capfd):
    # over stdio, stdout IS the protocol: the command's print() must be captured
    run(_call(srv.create_server(), "list_calendars"))
    assert capfd.readouterr().out == ""


# ------------------------------------------------------------------ writing

def test_validation_error_is_a_readable_tool_error_and_writes_nothing(graph):
    server = srv.create_server(allow_write=True)
    res = run(_call(server, "create_event",
                    {"subject": "X", "start": "2026-10-01T11:00:00", "end": "2026-10-01T10:00:00"}))
    assert res.is_error and "end" in _text(res).lower()
    assert not any(m == "POST" for m, _, _ in graph.calls)


def test_dry_run_builds_body_without_writing(graph):
    server = srv.create_server(allow_write=True)
    data = _data(run(_call(server, "create_event", {"subject": "Standup", "start": "2026-10-01T09:00:00",
                                                    "end": "2026-10-01T09:15:00", "dry_run": True})))
    assert data["dry_run"] is True and data["body"]["transactionId"]
    assert not any(m == "POST" for m, _, _ in graph.calls)


def test_create_event_writes_once_and_audits(graph):
    server = srv.create_server(allow_write=True)
    data = _data(run(_call(server, "create_event", {"subject": "Standup", "start": "2026-10-01T09:00:00",
                                                    "end": "2026-10-01T09:15:00", "location": "Teams"})))
    assert data["id"] == "EVNEW"
    posts = [b for m, _, b in graph.calls if m == "POST"]
    assert len(posts) == 1 and posts[0]["location"] == {"displayName": "Teams"}
    log = [json.loads(line) for line in o.AUDIT_LOG.read_text().splitlines()]
    assert log[-1]["op"] == "create-event" and log[-1]["id"] == "EVNEW"


def test_graph_error_comes_back_as_tool_error(graph):
    graph.fail = (403, "Access denied")
    res = run(_call(srv.create_server(), "list_calendars"))
    assert res.is_error and "403" in _text(res)


def test_unknown_calendar_name_is_explained(graph):
    res = run(_call(srv.create_server(), "list_events", {"calendar": "Gibtsnicht"}))
    assert res.is_error and "not found" in _text(res) and "Kalender" in _text(res)


# ------------------------------------------------- arguments and annotations

def test_argv_values_can_never_become_options():
    argv = srv.build_argv("create-task", title="-h", body="--dry-run", importance=None, dry_run=False)
    assert argv == ["create-task", "--title=-h", "--body=--dry-run"]
    assert srv.build_argv("complete-task", "-abc", task_list="Einkauf") == ["complete-task", "--list=Einkauf", "--", "-abc"]
    # "" is "not given" - except where it means something (update_task: clear the due date)
    assert srv.build_argv("update-task", "T1", title="") == ["update-task", "--", "T1"]
    assert srv.build_argv("update-task", "T1", keep_empty=("due",), due="") == ["update-task", "--due=", "--", "T1"]


def test_leading_dash_values_reach_graph_as_values(graph):
    server = srv.create_server(allow_write=True)
    data = _data(run(_call(server, "create_task", {"title": "-h", "body": "--dry-run"})))
    assert data["id"] == "TNEW"
    body = next(b for m, _, b in graph.calls if m == "POST")
    assert body["title"] == "-h" and body["body"]["content"] == "--dry-run"
    run(_call(server, "complete_task", {"task_id": "-weird-id"}))
    assert any(m == "PATCH" and p.endswith("/tasks/-weird-id") for m, p, _ in graph.calls)


def test_only_deduplicated_creates_claim_idempotency():
    tools = run(_tools(srv.create_server(allow_write=True)))
    assert tools["create_event"].annotations.idempotent_hint is True     # transactionId
    assert tools["create_task"].annotations.idempotent_hint is False     # a retry duplicates
    assert tools["add_step"].annotations.idempotent_hint is False


def test_unexpected_failure_is_named_not_a_crash(graph, monkeypatch):
    monkeypatch.setattr(o, "gapi", lambda *a, **k: {"value": [{"no": "id"}]})
    res = run(_call(srv.create_server(), "list_calendars"))
    assert res.is_error and "unexpected: KeyError" in _text(res)


# ------------------------------------------- IDs from a model are untrusted

@pytest.mark.parametrize("bad", ["../../me/messages", "AAMk?$select=x", "abc#frag", "a%2Fb", "x/y", "a b"])
def test_malformed_ids_are_refused_before_graph(graph, bad):
    res = run(_call(srv.create_server(), "get_event", {"event_id": bad}))
    assert res.is_error and "Graph ID" in _text(res)
    assert graph.calls == []


def test_write_and_delete_tools_check_ids_too(graph):
    server = srv.create_server(allow_write=True, allow_delete=True)
    for tool, args in (("update_event", {"event_id": "a/../b", "subject": "x"}),
                       ("complete_task", {"task_id": "t?x"}),
                       ("complete_step", {"task_id": "T1", "step_id": "s/../x"}),
                       ("delete_event", {"event_id": "e#x"}),
                       ("delete_task", {"task_id": "t%2F"})):
        res = run(_call(server, tool, args))
        assert res.is_error and "Graph ID" in _text(res), tool
    assert graph.calls == []


def test_real_graph_ids_pass():
    srv.check_id("event_id", "AAMkAGI2TG93AAA=")
    srv.check_id("task_id", "AQMkADAwATM0MDAAMS1hZjc2LTc3Y2UtMDACLTAwCgBGAAAD-_x+y")


def test_calendar_that_looks_like_an_id_must_be_clean_but_names_stay_free(graph):
    fake_id = "AAMk" + "A" * 60 + "/../../me"
    res = run(_call(srv.create_server(), "list_events", {"calendar": fake_id}))
    assert res.is_error and "Graph ID" in _text(res) and graph.calls == []
    res = run(_call(srv.create_server(), "list_events", {"calendar": "Arbeit/Privat"}))
    assert res.is_error and "not found" in _text(res)   # a plain name: looked up by name, never put into a URL


def test_instructions_name_the_time_zone_and_untrusted_text():
    text = srv.create_server().instructions
    assert "OUTLOOK_TZ" in text and srv.outlook.TZ in text and "data, not instructions" in text


# ----------------------------------------------------------- real transport

def test_starts_over_stdio_as_a_subprocess():
    params = StdioServerParameters(command=sys.executable, args=[str(ROOT / "scripts" / "mcp_server.py")],
                                   env={"M365_MCP_WRITE": "1"})

    async def go():
        async with Client(params) as c:
            return {t.name for t in (await c.list_tools()).tools}
    assert run(go()) == READ_TOOLS | WRITE_TOOLS
