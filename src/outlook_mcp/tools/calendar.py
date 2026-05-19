"""Calendar tools."""

from __future__ import annotations

from datetime import timedelta
from typing import Literal

from msgraph.generated.models.attendee import Attendee
from msgraph.generated.models.attendee_base import AttendeeBase
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.event import Event
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.location import Location
from msgraph.generated.models.time_constraint import TimeConstraint
from msgraph.generated.models.time_slot import TimeSlot
from msgraph.generated.users.item.calendars.item.calendar_view.calendar_view_request_builder import (
    CalendarViewRequestBuilder,
)
from msgraph.generated.users.item.events.item.accept.accept_post_request_body import (
    AcceptPostRequestBody,
)
from msgraph.generated.users.item.events.item.cancel.cancel_post_request_body import (
    CancelPostRequestBody,
)
from msgraph.generated.users.item.events.item.decline.decline_post_request_body import (
    DeclinePostRequestBody,
)
from msgraph.generated.users.item.events.item.tentatively_accept.tentatively_accept_post_request_body import (
    TentativelyAcceptPostRequestBody,
)
from msgraph.generated.users.item.find_meeting_times.find_meeting_times_post_request_body import (
    FindMeetingTimesPostRequestBody,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphValidationError, map_kiota_error
from outlook_mcp.graph.pagination import (
    decode_page_token,
    encode_next_link,
    validate_limit,
)
from outlook_mcp.graph.serialize import calendar_to_dict, event_to_dict
from outlook_mcp.graph.trimming import trim_calendar, trim_event


_VALID_RESPONSES = {
    "accept": "accepted",
    "tentativelyAccept": "tentativelyAccepted",
    "decline": "declined",
}


async def list_calendars(
    *, graph, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """List the user's calendars."""
    try:
        collection = await graph.mailbox(mailbox).calendars.get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [
        trim_calendar(calendar_to_dict(c), include_raw=include_raw)
        for c in (collection.value or [])
    ]
    return {"items": items, "next_page_token": None}


def _calendar_view_query(*, start_datetime: str, end_datetime: str, limit: int):
    qp = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
        start_date_time=start_datetime,
        end_date_time=end_datetime,
        top=limit,
        orderby=["start/dateTime"],
    )
    return CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
        query_parameters=qp
    )


async def list_events(
    *,
    graph,
    calendar_id: str,
    start_datetime: str | None = None,
    end_datetime: str | None = None,
    mailbox: str | None = None,
    limit: int = 25,
    page_token: str | None = None,
    include_raw: bool = False,
) -> dict:
    """List events in a calendar within a date range.

    Uses Graph's calendarView, which expands recurring events into instances.

    Args:
        calendar_id: Calendar ID. (Use list_calendars to discover.)
        start_datetime: ISO 8601 datetime, e.g. "2026-05-19T00:00:00Z". Required.
        end_datetime: ISO 8601 datetime. Required.
        mailbox: Optional mailbox.
        limit: 1-100. Default 25.
        page_token: Continuation token.
        include_raw: Include raw payloads.

    Returns:
        {"items": [trimmed_event, ...], "next_page_token": str | None}
    """
    if not (start_datetime and end_datetime):
        raise GraphValidationError("start_datetime and end_datetime are required")
    limit = validate_limit(limit)
    cal_builder = graph.mailbox(mailbox).calendars.by_calendar_id(calendar_id)
    try:
        if page_token is not None:
            url = decode_page_token(page_token)
            collection = await cal_builder.calendar_view.with_url(url).get()
        else:
            collection = await cal_builder.calendar_view.get(
                request_configuration=_calendar_view_query(
                    start_datetime=start_datetime, end_datetime=end_datetime, limit=limit
                )
            )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    items = [
        trim_event(event_to_dict(e), include_body=False, include_raw=include_raw)
        for e in (collection.value or [])
    ]
    return {
        "items": items,
        "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None)),
    }


async def get_event(
    *,
    graph,
    event_id: str,
    mailbox: str | None = None,
    include_body: bool = False,
    include_raw: bool = False,
) -> dict:
    """Fetch a single event by id."""
    try:
        evt = await graph.mailbox(mailbox).events.by_event_id(event_id).get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_event(event_to_dict(evt), include_body=include_body, include_raw=include_raw)


def _dttz(dt: str, tz: str) -> DateTimeTimeZone:
    obj = DateTimeTimeZone()
    obj.date_time = dt
    obj.time_zone = tz
    return obj


def _build_event(
    *,
    subject: str,
    start_datetime: str,
    end_datetime: str,
    time_zone: str,
    body: str | None,
    body_type: Literal["text", "html"],
    location: str | None,
    attendees: list[str] | None,
    is_online_meeting: bool,
    is_all_day: bool,
) -> Event:
    e = Event()
    e.subject = subject
    e.start = _dttz(start_datetime, time_zone)
    e.end = _dttz(end_datetime, time_zone)
    if body is not None:
        b = ItemBody()
        b.content_type = BodyType.Text if body_type == "text" else BodyType.Html
        b.content = body
        e.body = b
    if location is not None:
        loc = Location()
        loc.display_name = location
        e.location = loc
    if attendees:
        atts: list[Attendee] = []
        for addr in attendees:
            a = Attendee()
            ea = EmailAddress()
            ea.address = addr
            a.email_address = ea
            a.type = AttendeeType.Required
            atts.append(a)
        e.attendees = atts
    e.is_online_meeting = is_online_meeting
    e.is_all_day = is_all_day
    return e


async def create_event(
    *,
    graph,
    subject: str,
    start_datetime: str,
    end_datetime: str,
    time_zone: str = "UTC",
    body: str | None = None,
    body_type: Literal["text", "html"] = "text",
    location: str | None = None,
    attendees: list[str] | None = None,
    is_online_meeting: bool = False,
    is_all_day: bool = False,
    calendar_id: str | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Create a calendar event.

    Args:
        subject: Event subject.
        start_datetime: ISO 8601 datetime (no offset; pair with time_zone).
        end_datetime: ISO 8601 datetime.
        time_zone: IANA tz id (e.g. "America/Los_Angeles") or "UTC". Default "UTC".
        body, body_type: Optional body and "text"|"html".
        location: Optional location string.
        attendees: Optional list of attendee email addresses (treated as required).
        is_online_meeting: When True, Outlook adds a Teams meeting link.
        is_all_day: All-day event flag.
        calendar_id: Optional calendar id; default is the user's primary calendar.
        mailbox: Optional mailbox.

    Returns:
        Trimmed created event.
    """
    evt = _build_event(
        subject=subject, start_datetime=start_datetime, end_datetime=end_datetime,
        time_zone=time_zone, body=body, body_type=body_type, location=location,
        attendees=attendees, is_online_meeting=is_online_meeting, is_all_day=is_all_day,
    )
    try:
        target = graph.mailbox(mailbox)
        if calendar_id is not None:
            target = target.calendars.by_calendar_id(calendar_id)
        created = await target.events.post(evt)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_event(event_to_dict(created), include_body=False, include_raw=include_raw)


async def update_event(
    *,
    graph,
    event_id: str,
    subject: str | None = None,
    start_datetime: str | None = None,
    end_datetime: str | None = None,
    time_zone: str | None = None,
    body: str | None = None,
    body_type: Literal["text", "html"] = "text",
    location: str | None = None,
    is_online_meeting: bool | None = None,
    is_all_day: bool | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Patch fields of an existing event. Pass None to leave a field unchanged.

    If you change start_datetime or end_datetime, you must also pass time_zone.
    """
    if (start_datetime or end_datetime) and not time_zone:
        raise GraphValidationError("time_zone is required when updating start or end")
    patch = Event()
    if subject is not None:
        patch.subject = subject
    if start_datetime is not None:
        patch.start = _dttz(start_datetime, time_zone or "UTC")
    if end_datetime is not None:
        patch.end = _dttz(end_datetime, time_zone or "UTC")
    if body is not None:
        b = ItemBody()
        b.content_type = BodyType.Text if body_type == "text" else BodyType.Html
        b.content = body
        patch.body = b
    if location is not None:
        loc = Location()
        loc.display_name = location
        patch.location = loc
    if is_online_meeting is not None:
        patch.is_online_meeting = is_online_meeting
    if is_all_day is not None:
        patch.is_all_day = is_all_day
    try:
        updated = await graph.mailbox(mailbox).events.by_event_id(event_id).patch(patch)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_event(event_to_dict(updated), include_body=False, include_raw=include_raw)


async def delete_event(*, graph, event_id: str, mailbox: str | None = None) -> dict:
    """Delete an event without sending a cancellation to attendees."""
    try:
        await graph.mailbox(mailbox).events.by_event_id(event_id).delete()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "deleted"}


async def cancel_event(
    *, graph, event_id: str, comment: str | None = None, mailbox: str | None = None
) -> dict:
    """Cancel an event and send a cancellation notice to attendees."""
    body = CancelPostRequestBody()
    if comment is not None:
        body.comment = comment
    try:
        await graph.mailbox(mailbox).events.by_event_id(event_id).cancel.post(body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "cancelled"}


async def respond_to_event(
    *,
    graph,
    event_id: str,
    response: Literal["accept", "tentativelyAccept", "decline"],
    comment: str | None = None,
    send_response: bool = True,
    mailbox: str | None = None,
) -> dict:
    """Respond to a meeting invite.

    Args:
        response: One of "accept", "tentativelyAccept", "decline".
        comment: Optional comment included with the response.
        send_response: When False, your response status is recorded without
            emailing the organizer.
    """
    if response not in _VALID_RESPONSES:
        raise GraphValidationError(
            f"response must be one of {sorted(_VALID_RESPONSES)}"
        )
    evt_builder = graph.mailbox(mailbox).events.by_event_id(event_id)
    try:
        if response == "accept":
            accept_body = AcceptPostRequestBody()
            accept_body.comment = comment
            accept_body.send_response = send_response
            await evt_builder.accept.post(accept_body)
        elif response == "tentativelyAccept":
            tentative_body = TentativelyAcceptPostRequestBody()
            tentative_body.comment = comment
            tentative_body.send_response = send_response
            await evt_builder.tentatively_accept.post(tentative_body)
        else:  # decline
            decline_body = DeclinePostRequestBody()
            decline_body.comment = comment
            decline_body.send_response = send_response
            await evt_builder.decline.post(decline_body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": _VALID_RESPONSES[response]}


def _duration_td(minutes: int) -> timedelta:
    """Return a timedelta for the given whole minutes.

    The kiota serializer renders this as ISO 8601 (e.g. PT30M) on the wire.
    """
    if minutes <= 0:
        raise GraphValidationError("duration_minutes must be a positive integer")
    return timedelta(minutes=int(minutes))


async def find_meeting_times(
    *,
    graph,
    attendees: list[str],
    duration_minutes: int,
    start_window: str,
    end_window: str,
    mailbox: str | None = None,
    max_candidates: int = 20,
    include_raw: bool = False,
) -> dict:
    """Suggest meeting times that work for the given attendees.

    Args:
        attendees: List of attendee email addresses.
        duration_minutes: Meeting duration in whole minutes (e.g. 30).
        start_window: ISO 8601 start of the candidate window.
        end_window: ISO 8601 end of the candidate window.
        max_candidates: Cap on returned suggestions (default 20).

    Returns:
        {
          "suggestions": [
            {"start": {date_time, time_zone}, "end": {...}, "confidence": float, "order_hint": int},
            ...
          ],
          "empty_reason": str | None,
        }
    """
    if not attendees:
        raise GraphValidationError("`attendees` must be a non-empty list")
    body = FindMeetingTimesPostRequestBody()
    body.meeting_duration = _duration_td(duration_minutes)
    body.max_candidates = max_candidates
    body.attendees = []
    for addr in attendees:
        a = AttendeeBase()
        ea = EmailAddress()
        ea.address = addr
        a.email_address = ea
        body.attendees.append(a)
    tc = TimeConstraint()
    slot = TimeSlot()
    slot.start = _dttz(start_window, "UTC")
    slot.end = _dttz(end_window, "UTC")
    tc.time_slots = [slot]
    body.time_constraint = tc
    try:
        result = await graph.mailbox(mailbox).find_meeting_times.post(body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    suggestions_raw = getattr(result, "meeting_time_suggestions", None) or []
    suggestions = []
    for s in suggestions_raw:
        ts = getattr(s, "meeting_time_slot", None)
        suggestions.append(
            {
                "start": {
                    "date_time": getattr(getattr(ts, "start", None), "date_time", None),
                    "time_zone": getattr(getattr(ts, "start", None), "time_zone", None),
                } if ts else None,
                "end": {
                    "date_time": getattr(getattr(ts, "end", None), "date_time", None),
                    "time_zone": getattr(getattr(ts, "end", None), "time_zone", None),
                } if ts else None,
                "confidence": getattr(s, "confidence", None),
                "order_hint": getattr(s, "order_hint", None),
            }
        )
    out: dict = {
        "suggestions": suggestions,
        "empty_reason": getattr(result, "empty_suggestions_reason", None),
    }
    if include_raw:
        out["raw"] = getattr(result, "additional_data", {}) or {}
    return out


def register(mcp, *, graph) -> None:
    @mcp.tool(name="list_calendars", description=list_calendars.__doc__ or "")
    async def _list_calendars(mailbox: str | None = None, include_raw: bool = False):
        return await list_calendars(graph=graph, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="list_events", description=list_events.__doc__ or "")
    async def _list_events(
        calendar_id: str,
        start_datetime: str,
        end_datetime: str,
        mailbox: str | None = None,
        limit: int = 25,
        page_token: str | None = None,
        include_raw: bool = False,
    ):
        return await list_events(
            graph=graph, calendar_id=calendar_id,
            start_datetime=start_datetime, end_datetime=end_datetime,
            mailbox=mailbox, limit=limit, page_token=page_token, include_raw=include_raw,
        )

    @mcp.tool(name="get_event", description=get_event.__doc__ or "")
    async def _get_event(
        event_id: str,
        mailbox: str | None = None,
        include_body: bool = False,
        include_raw: bool = False,
    ):
        return await get_event(
            graph=graph, event_id=event_id, mailbox=mailbox,
            include_body=include_body, include_raw=include_raw,
        )

    @mcp.tool(name="create_event", description=create_event.__doc__ or "")
    async def _create_event(
        subject: str, start_datetime: str, end_datetime: str,
        time_zone: str = "UTC",
        body: str | None = None,
        body_type: Literal["text", "html"] = "text",
        location: str | None = None, attendees: list[str] | None = None,
        is_online_meeting: bool = False, is_all_day: bool = False,
        calendar_id: str | None = None, mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await create_event(
            graph=graph, subject=subject,
            start_datetime=start_datetime, end_datetime=end_datetime,
            time_zone=time_zone, body=body, body_type=body_type, location=location,
            attendees=attendees, is_online_meeting=is_online_meeting, is_all_day=is_all_day,
            calendar_id=calendar_id, mailbox=mailbox, include_raw=include_raw,
        )

    @mcp.tool(name="update_event", description=update_event.__doc__ or "")
    async def _update_event(
        event_id: str,
        subject: str | None = None,
        start_datetime: str | None = None,
        end_datetime: str | None = None,
        time_zone: str | None = None,
        body: str | None = None,
        body_type: Literal["text", "html"] = "text",
        location: str | None = None,
        is_online_meeting: bool | None = None,
        is_all_day: bool | None = None,
        mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await update_event(
            graph=graph, event_id=event_id, subject=subject,
            start_datetime=start_datetime, end_datetime=end_datetime, time_zone=time_zone,
            body=body, body_type=body_type, location=location,
            is_online_meeting=is_online_meeting, is_all_day=is_all_day,
            mailbox=mailbox, include_raw=include_raw,
        )

    @mcp.tool(name="delete_event", description=delete_event.__doc__ or "")
    async def _delete_event(event_id: str, mailbox: str | None = None):
        return await delete_event(graph=graph, event_id=event_id, mailbox=mailbox)

    @mcp.tool(name="cancel_event", description=cancel_event.__doc__ or "")
    async def _cancel_event(
        event_id: str, comment: str | None = None, mailbox: str | None = None
    ):
        return await cancel_event(
            graph=graph, event_id=event_id, comment=comment, mailbox=mailbox
        )

    @mcp.tool(name="respond_to_event", description=respond_to_event.__doc__ or "")
    async def _respond(
        event_id: str,
        response: Literal["accept", "tentativelyAccept", "decline"],
        comment: str | None = None,
        send_response: bool = True,
        mailbox: str | None = None,
    ):
        return await respond_to_event(
            graph=graph, event_id=event_id, response=response, comment=comment,
            send_response=send_response, mailbox=mailbox,
        )

    @mcp.tool(name="find_meeting_times", description=find_meeting_times.__doc__ or "")
    async def _find(
        attendees: list[str], duration_minutes: int,
        start_window: str, end_window: str,
        mailbox: str | None = None, max_candidates: int = 20,
        include_raw: bool = False,
    ):
        return await find_meeting_times(
            graph=graph, attendees=attendees, duration_minutes=duration_minutes,
            start_window=start_window, end_window=end_window,
            mailbox=mailbox, max_candidates=max_candidates, include_raw=include_raw,
        )
