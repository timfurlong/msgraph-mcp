from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import calendar as cal_tools


def _fake_calendar(id_="c1", name="Calendar"):
    return SimpleNamespace(
        id=id_, name=name,
        owner=SimpleNamespace(name="Alice", address="alice@example.com"),
        can_edit=True, is_default_calendar=True, additional_data={},
    )


def _fake_event(id_="e1"):
    return SimpleNamespace(
        id=id_,
        subject="Standup",
        organizer=SimpleNamespace(
            email_address=SimpleNamespace(name="Alice", address="alice@example.com")
        ),
        start=SimpleNamespace(date_time="2026-05-19T15:00:00", time_zone="UTC"),
        end=SimpleNamespace(date_time="2026-05-19T15:30:00", time_zone="UTC"),
        location=SimpleNamespace(display_name="Zoom"),
        is_all_day=False, is_cancelled=False, is_online_meeting=False, online_meeting=None,
        attendees=[],
        body_preview="Daily standup",
        body=SimpleNamespace(content_type=SimpleNamespace(value="text"), content="Daily standup"),
        show_as=SimpleNamespace(value="busy"),
        sensitivity=SimpleNamespace(value="normal"),
        recurrence=None,
        web_link="https://outlook.example/evt",
        additional_data={},
    )


def _fake_collection(items, next_link=None):
    return SimpleNamespace(value=items, odata_next_link=next_link)


@pytest.mark.asyncio
async def test_list_calendars():
    graph = MagicMock()
    mb = MagicMock()
    mb.calendars.get = AsyncMock(return_value=_fake_collection([_fake_calendar()]))
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.list_calendars(graph=graph)
    assert result["items"][0]["name"] == "Calendar"


@pytest.mark.asyncio
async def test_list_events_uses_calendar_view_with_date_range():
    graph = MagicMock()
    mb = MagicMock()
    cal_builder = MagicMock()
    cal_builder.calendar_view.get = AsyncMock(return_value=_fake_collection([_fake_event()]))
    mb.calendars.by_calendar_id = MagicMock(return_value=cal_builder)
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.list_events(
        graph=graph,
        calendar_id="c1",
        start_datetime="2026-05-19T00:00:00Z",
        end_datetime="2026-05-20T00:00:00Z",
    )
    cal_builder.calendar_view.get.assert_awaited_once()
    assert result["items"][0]["id"] == "e1"


@pytest.mark.asyncio
async def test_list_events_requires_both_datetimes():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await cal_tools.list_events(
            graph=graph, calendar_id="c1", start_datetime="2026-05-19T00:00:00Z"
        )


@pytest.mark.asyncio
async def test_get_event_returns_trimmed():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.by_event_id = MagicMock(
        return_value=MagicMock(get=AsyncMock(return_value=_fake_event()))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.get_event(graph=graph, event_id="e1")
    assert result["id"] == "e1"


@pytest.mark.asyncio
async def test_create_event_posts_to_events():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.post = AsyncMock(return_value=_fake_event(id_="new1"))
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.create_event(
        graph=graph,
        subject="Standup",
        start_datetime="2026-05-19T15:00:00",
        end_datetime="2026-05-19T15:30:00",
        time_zone="UTC",
    )
    mb.events.post.assert_awaited_once()
    assert result["id"] == "new1"


@pytest.mark.asyncio
async def test_update_event_patches():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.by_event_id = MagicMock(
        return_value=MagicMock(patch=AsyncMock(return_value=_fake_event(id_="e1")))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.update_event(graph=graph, event_id="e1", subject="Updated")
    assert result["id"] == "e1"


@pytest.mark.asyncio
async def test_delete_event_calls_delete():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.by_event_id = MagicMock(
        return_value=MagicMock(delete=AsyncMock(return_value=None))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.delete_event(graph=graph, event_id="e1")
    assert result == {"status": "deleted"}


@pytest.mark.asyncio
async def test_cancel_event_calls_cancel_endpoint():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.by_event_id = MagicMock(
        return_value=MagicMock(cancel=MagicMock(post=AsyncMock(return_value=None)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.cancel_event(
        graph=graph, event_id="e1", comment="Sorry, conflict."
    )
    assert result == {"status": "cancelled"}


@pytest.mark.asyncio
async def test_respond_to_event_accept():
    graph = MagicMock()
    mb = MagicMock()
    mb.events.by_event_id = MagicMock(
        return_value=MagicMock(accept=MagicMock(post=AsyncMock(return_value=None)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.respond_to_event(graph=graph, event_id="e1", response="accept")
    assert result == {"status": "accepted"}


@pytest.mark.asyncio
async def test_respond_to_event_rejects_bad_response():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        # Deliberately invalid response value to exercise the runtime guard.
        await cal_tools.respond_to_event(graph=graph, event_id="e1", response="maybe")  # type: ignore[arg-type]


@pytest.mark.asyncio
async def test_find_meeting_times_returns_suggestions():
    graph = MagicMock()
    mb = MagicMock()
    suggestion = SimpleNamespace(
        meeting_time_slot=SimpleNamespace(
            start=SimpleNamespace(date_time="2026-05-20T10:00:00", time_zone="UTC"),
            end=SimpleNamespace(date_time="2026-05-20T10:30:00", time_zone="UTC"),
        ),
        confidence=85.0,
        order_hint=1,
    )
    result_obj = SimpleNamespace(
        meeting_time_suggestions=[suggestion],
        empty_suggestions_reason=None,
        additional_data={},
    )
    mb.find_meeting_times.post = AsyncMock(return_value=result_obj)
    graph.mailbox = MagicMock(return_value=mb)

    result = await cal_tools.find_meeting_times(
        graph=graph,
        attendees=["bob@example.com"],
        duration_minutes=30,
        start_window="2026-05-20T00:00:00Z",
        end_window="2026-05-21T00:00:00Z",
    )
    assert len(result["suggestions"]) == 1
    assert result["suggestions"][0]["confidence"] == 85.0
