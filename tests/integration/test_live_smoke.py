"""Live integration smoke tests against the real mailbox + calendar.

Gated behind OUTLOOK_MCP_INTEGRATION=1 via tests/conftest.py.

These tests require a valid token cache at ~/.outlook-mcp/token_cache.bin
(run `outlook-mcp-login` first). They make real Graph calls and may
create/delete real items in your mailbox and a temporary calendar.
"""

from __future__ import annotations

import datetime as dt
import time

import pytest

from outlook_mcp.graph.client import GraphClient
from outlook_mcp.tools import calendar as cal_tools
from outlook_mcp.tools import mail_read, mail_write
from outlook_mcp.tools import util as util_tools


@pytest.fixture
def graph():
    # Function scope — pytest-asyncio creates a fresh event loop per test, and
    # the kiota httpx client inside GraphServiceClient is bound to the loop
    # alive at construction time. A module-scoped fixture would reuse a client
    # whose loop has been closed by the first test.
    return GraphClient()


@pytest.mark.asyncio
async def test_whoami(graph):
    me = await util_tools.whoami(graph=graph)
    assert me["mail"], "Expected /me to return a mail address"


@pytest.mark.asyncio
async def test_list_messages_returns_at_least_zero(graph):
    result = await mail_read.list_messages(graph=graph, limit=1)
    assert "items" in result
    assert isinstance(result["items"], list)


@pytest.mark.asyncio
async def test_list_calendars(graph):
    result = await cal_tools.list_calendars(graph=graph)
    assert len(result["items"]) >= 1


@pytest.mark.asyncio
async def test_event_create_get_delete_roundtrip(graph):
    # Use the user's primary calendar.
    cals = await cal_tools.list_calendars(graph=graph)
    default_cal = next(
        (c for c in cals["items"] if c["is_default_calendar"]),
        cals["items"][0],
    )

    start = dt.datetime.now(dt.UTC).replace(microsecond=0, tzinfo=None) + dt.timedelta(days=365)
    end = start + dt.timedelta(minutes=15)
    subject = f"Outlook MCP smoke {int(time.time())}"

    created = await cal_tools.create_event(
        graph=graph,
        subject=subject,
        start_datetime=start.isoformat(),
        end_datetime=end.isoformat(),
        time_zone="UTC",
        body="Created by tests/integration/test_live_smoke.py",
        calendar_id=default_cal["id"],
    )
    event_id = created["id"]
    try:
        fetched = await cal_tools.get_event(graph=graph, event_id=event_id)
        assert fetched["subject"] == subject
    finally:
        result = await cal_tools.delete_event(graph=graph, event_id=event_id)
        assert result == {"status": "deleted"}


@pytest.mark.asyncio
async def test_send_to_self_search_and_delete(graph):
    me = await util_tools.whoami(graph=graph)
    # search_messages wraps the query in outer double quotes, so avoid inner
    # quotes/KQL keywords. A unique token in the subject is enough to find the
    # message via plain $search.
    token = f"MCPSMOKE{int(time.time())}"
    subject = f"Outlook MCP smoke msg {token}"
    sent = await mail_write.send_message(
        graph=graph,
        to=[me["mail"]],
        subject=subject,
        body="Sent by tests/integration/test_live_smoke.py — safe to delete",
    )
    assert sent == {"status": "sent"}
    # Allow Graph indexing latency.
    found = None
    for _ in range(10):
        result = await mail_read.search_messages(graph=graph, query=token, limit=5)
        if result["items"]:
            found = result["items"][0]
            break
        time.sleep(2)
    assert found is not None, "Sent message never appeared in search results"
    # Cleanup
    deleted = await mail_write.delete_message(graph=graph, message_id=found["id"])
    assert deleted == {"status": "deleted"}
