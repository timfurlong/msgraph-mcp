"""Trim raw Graph dicts to small agent-friendly shapes.

All trimmers accept dict-shaped input (raw Graph JSON), not msgraph-sdk
model objects. Tool implementations are responsible for calling
.serialize() or otherwise producing a plain dict before trimming.

include_raw=True returns {**trimmed, "raw": raw}.
"""

from __future__ import annotations

from typing import Any


def _attach_raw(trimmed: dict, raw: dict, include_raw: bool) -> dict:
    if include_raw:
        return {**trimmed, "raw": raw}
    return trimmed


def _email(rec: dict | None) -> dict | None:
    if not rec:
        return None
    addr = rec.get("emailAddress") or {}
    return {"name": addr.get("name"), "address": addr.get("address")}


def _emails(recs: list[dict] | None) -> list[dict]:
    return [e for e in (_email(r) for r in (recs or [])) if e is not None]


def trim_message(raw: dict, *, include_body: bool, include_raw: bool) -> dict:
    body = raw.get("body") or {}
    flag = raw.get("flag") or {}
    trimmed: dict[str, Any] = {
        "id": raw.get("id"),
        "subject": raw.get("subject"),
        "from": _email(raw.get("from")),
        "to": _emails(raw.get("toRecipients")),
        "cc": _emails(raw.get("ccRecipients")),
        "received": raw.get("receivedDateTime"),
        "sent": raw.get("sentDateTime"),
        "snippet": raw.get("bodyPreview"),
        "body_type": body.get("contentType") or "text",
        "is_read": bool(raw.get("isRead")),
        "is_draft": bool(raw.get("isDraft")),
        "has_attachments": bool(raw.get("hasAttachments")),
        "importance": raw.get("importance") or "normal",
        "flag": flag.get("flagStatus") or "notFlagged",
        "categories": list(raw.get("categories") or []),
        "conversation_id": raw.get("conversationId"),
        "web_link": raw.get("webLink"),
    }
    if include_body:
        trimmed["body"] = body.get("content")
    return _attach_raw(trimmed, raw, include_raw)


def trim_event(raw: dict, *, include_body: bool, include_raw: bool) -> dict:
    start = raw.get("start") or {}
    end = raw.get("end") or {}
    location = raw.get("location") or {}
    online = raw.get("onlineMeeting") or {}
    body = raw.get("body") or {}
    attendees_raw = raw.get("attendees") or []
    attendees = []
    for a in attendees_raw:
        email = _email(a)
        status = (a.get("status") or {}).get("response")
        attendees.append(
            {
                "name": (email or {}).get("name"),
                "address": (email or {}).get("address"),
                "type": a.get("type"),
                "response": status,
            }
        )
    trimmed: dict[str, Any] = {
        "id": raw.get("id"),
        "subject": raw.get("subject"),
        "organizer": _email(raw.get("organizer")),
        "start": {"date_time": start.get("dateTime"), "time_zone": start.get("timeZone")},
        "end": {"date_time": end.get("dateTime"), "time_zone": end.get("timeZone")},
        "location": location.get("displayName"),
        "is_all_day": bool(raw.get("isAllDay")),
        "is_cancelled": bool(raw.get("isCancelled")),
        "is_online_meeting": bool(raw.get("isOnlineMeeting")),
        "attendees": attendees,
        "body_preview": raw.get("bodyPreview"),
        "show_as": raw.get("showAs"),
        "sensitivity": raw.get("sensitivity"),
        "recurrence": raw.get("recurrence"),
        "web_link": raw.get("webLink"),
    }
    if raw.get("isOnlineMeeting"):
        trimmed["online_meeting_url"] = online.get("joinUrl")
    if include_body:
        trimmed["body"] = body.get("content")
    return _attach_raw(trimmed, raw, include_raw)


def trim_folder(raw: dict, *, include_raw: bool) -> dict:
    trimmed = {
        "id": raw.get("id"),
        "display_name": raw.get("displayName"),
        "parent_folder_id": raw.get("parentFolderId"),
        "total_item_count": raw.get("totalItemCount"),
        "unread_item_count": raw.get("unreadItemCount"),
        "child_folder_count": raw.get("childFolderCount"),
    }
    return _attach_raw(trimmed, raw, include_raw)


def trim_attachment_list(raw: dict, *, include_raw: bool) -> dict:
    trimmed = {
        "id": raw.get("id"),
        "name": raw.get("name"),
        "content_type": raw.get("contentType"),
        "size_bytes": raw.get("size"),
        "is_inline": bool(raw.get("isInline")),
    }
    return _attach_raw(trimmed, raw, include_raw)


def trim_attachment_download(raw: dict, *, include_raw: bool) -> dict:
    trimmed = {
        "name": raw.get("name"),
        "content_type": raw.get("contentType"),
        "size_bytes": raw.get("size"),
        "content_base64": raw.get("contentBytes"),
    }
    return _attach_raw(trimmed, raw, include_raw)


def trim_user(raw: dict, *, include_raw: bool) -> dict:
    trimmed = {
        "id": raw.get("id"),
        "display_name": raw.get("displayName"),
        "user_principal_name": raw.get("userPrincipalName"),
        "mail": raw.get("mail"),
        "job_title": raw.get("jobTitle"),
    }
    return _attach_raw(trimmed, raw, include_raw)


def trim_calendar(raw: dict, *, include_raw: bool) -> dict:
    owner = raw.get("owner") or {}
    trimmed = {
        "id": raw.get("id"),
        "name": raw.get("name"),
        "owner": {"name": owner.get("name"), "address": owner.get("address")} if owner else None,
        "can_edit": bool(raw.get("canEdit")),
        "is_default_calendar": bool(raw.get("isDefaultCalendar")),
    }
    return _attach_raw(trimmed, raw, include_raw)


def _trim_rule_conditions(raw: dict | None) -> dict | None:
    if not raw:
        return None
    return {
        "sender_contains": raw.get("senderContains"),
        "subject_contains": raw.get("subjectContains"),
        "body_contains": raw.get("bodyContains"),
        "body_or_subject_contains": raw.get("bodyOrSubjectContains"),
        "from_addresses": _emails(raw.get("fromAddresses")),
        "has_attachments": raw.get("hasAttachments"),
        "categories": raw.get("categories"),
    }


def _trim_rule_actions(raw: dict | None) -> dict | None:
    if not raw:
        return None
    return {
        "move_to_folder": raw.get("moveToFolder"),
        "copy_to_folder": raw.get("copyToFolder"),
        "delete": raw.get("delete"),
        "mark_as_read": raw.get("markAsRead"),
        "mark_importance": raw.get("markImportance"),
        "assign_categories": raw.get("assignCategories"),
        "forward_to": _emails(raw.get("forwardTo")),
        "stop_processing_rules": raw.get("stopProcessingRules"),
    }


def trim_message_rule(raw: dict, *, include_raw: bool) -> dict:
    trimmed = {
        "id": raw.get("id"),
        "display_name": raw.get("displayName"),
        "sequence": raw.get("sequence"),
        "is_enabled": bool(raw.get("isEnabled")),
        "has_error": bool(raw.get("hasError")),
        "is_read_only": bool(raw.get("isReadOnly")),
        "conditions": _trim_rule_conditions(raw.get("conditions")),
        "actions": _trim_rule_actions(raw.get("actions")),
    }
    return _attach_raw(trimmed, raw, include_raw)
