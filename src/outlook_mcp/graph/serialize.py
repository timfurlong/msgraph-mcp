"""Convert msgraph-sdk model objects into camelCase dicts.

The trimmers in graph.trimming consume dicts shaped like raw Graph JSON
(camelCase, nested objects). msgraph-sdk's deserialized models use
snake_case attrs and enum wrappers. These helpers bridge the two.

Pattern: each helper pulls the attrs we care about into a dict, then
merges with model.additional_data so include_raw=True returns whatever
Graph sent that the SDK didn't model.
"""

from __future__ import annotations

import base64
from typing import Any


def _enum_value(x: Any) -> Any:
    """Return .value if x looks like an enum-wrapper, else x."""
    return getattr(x, "value", x)


def _additional(obj: Any) -> dict:
    return getattr(obj, "additional_data", None) or {}


def user_to_dict(user: Any) -> dict:
    base = {
        "id": getattr(user, "id", None),
        "displayName": getattr(user, "display_name", None),
        "userPrincipalName": getattr(user, "user_principal_name", None),
        "mail": getattr(user, "mail", None),
        "jobTitle": getattr(user, "job_title", None),
    }
    return {**_additional(user), **base}


def email_address_to_dict(addr: Any) -> dict | None:
    if addr is None:
        return None
    return {"name": getattr(addr, "name", None), "address": getattr(addr, "address", None)}


def recipient_to_dict(rec: Any) -> dict | None:
    if rec is None:
        return None
    return {"emailAddress": email_address_to_dict(getattr(rec, "email_address", None))}


def _recipients(recs: list[Any] | None) -> list[dict]:
    return [r for r in (recipient_to_dict(rec) for rec in (recs or [])) if r is not None]


def _body_to_dict(body: Any) -> dict | None:
    if body is None:
        return None
    return {
        "contentType": _enum_value(getattr(body, "content_type", None)),
        "content": getattr(body, "content", None),
    }


def _flag_to_dict(flag: Any) -> dict | None:
    if flag is None:
        return None
    return {"flagStatus": _enum_value(getattr(flag, "flag_status", None))}


def message_to_dict(msg: Any) -> dict:
    base = {
        "id": getattr(msg, "id", None),
        "subject": getattr(msg, "subject", None),
        "from": recipient_to_dict(getattr(msg, "from_", None) or getattr(msg, "sender", None)),
        "toRecipients": _recipients(getattr(msg, "to_recipients", None)),
        "ccRecipients": _recipients(getattr(msg, "cc_recipients", None)),
        "bccRecipients": _recipients(getattr(msg, "bcc_recipients", None)),
        "receivedDateTime": getattr(msg, "received_date_time", None),
        "sentDateTime": getattr(msg, "sent_date_time", None),
        "bodyPreview": getattr(msg, "body_preview", None),
        "body": _body_to_dict(getattr(msg, "body", None)),
        "isRead": bool(getattr(msg, "is_read", False)),
        "isDraft": bool(getattr(msg, "is_draft", False)),
        "hasAttachments": bool(getattr(msg, "has_attachments", False)),
        "importance": _enum_value(getattr(msg, "importance", None)),
        "flag": _flag_to_dict(getattr(msg, "flag", None)),
        "categories": list(getattr(msg, "categories", None) or []),
        "conversationId": getattr(msg, "conversation_id", None),
        "webLink": getattr(msg, "web_link", None),
        "parentFolderId": getattr(msg, "parent_folder_id", None),
        "inferenceClassification": _enum_value(getattr(msg, "inference_classification", None)),
    }
    return {**_additional(msg), **base}


def folder_to_dict(folder: Any) -> dict:
    base = {
        "id": getattr(folder, "id", None),
        "displayName": getattr(folder, "display_name", None),
        "parentFolderId": getattr(folder, "parent_folder_id", None),
        "totalItemCount": getattr(folder, "total_item_count", None),
        "unreadItemCount": getattr(folder, "unread_item_count", None),
        "childFolderCount": getattr(folder, "child_folder_count", None),
    }
    return {**_additional(folder), **base}


def attachment_to_dict(att: Any, *, include_content: bool = False) -> dict:
    base = {
        "id": getattr(att, "id", None),
        "name": getattr(att, "name", None),
        "contentType": getattr(att, "content_type", None),
        "size": getattr(att, "size", None),
        "isInline": bool(getattr(att, "is_inline", False)),
    }
    if include_content:
        raw = getattr(att, "content_bytes", None)
        if isinstance(raw, (bytes, bytearray)):
            base["contentBytes"] = base64.b64encode(raw).decode("ascii")
        else:
            base["contentBytes"] = raw  # already a str
    return {**_additional(att), **base}


def calendar_to_dict(cal: Any) -> dict:
    owner = getattr(cal, "owner", None)
    base = {
        "id": getattr(cal, "id", None),
        "name": getattr(cal, "name", None),
        "owner": email_address_to_dict(owner) if owner is not None else None,
        "canEdit": bool(getattr(cal, "can_edit", False)),
        "isDefaultCalendar": bool(getattr(cal, "is_default_calendar", False)),
    }
    return {**_additional(cal), **base}


def event_to_dict(evt: Any) -> dict:
    start = getattr(evt, "start", None)
    end = getattr(evt, "end", None)
    location = getattr(evt, "location", None)
    online = getattr(evt, "online_meeting", None)
    attendees_raw = getattr(evt, "attendees", None) or []
    attendees = []
    for a in attendees_raw:
        status = getattr(a, "status", None)
        attendees.append(
            {
                "emailAddress": email_address_to_dict(getattr(a, "email_address", None)),
                "type": _enum_value(getattr(a, "type", None)),
                "status": {"response": _enum_value(getattr(status, "response", None))}
                if status is not None
                else None,
            }
        )
    base = {
        "id": getattr(evt, "id", None),
        "subject": getattr(evt, "subject", None),
        "organizer": recipient_to_dict(getattr(evt, "organizer", None)),
        "start": {
            "dateTime": getattr(start, "date_time", None) if start else None,
            "timeZone": getattr(start, "time_zone", None) if start else None,
        },
        "end": {
            "dateTime": getattr(end, "date_time", None) if end else None,
            "timeZone": getattr(end, "time_zone", None) if end else None,
        },
        "location": {"displayName": getattr(location, "display_name", None)} if location else {"displayName": None},
        "isAllDay": bool(getattr(evt, "is_all_day", False)),
        "isCancelled": bool(getattr(evt, "is_cancelled", False)),
        "isOnlineMeeting": bool(getattr(evt, "is_online_meeting", False)),
        "onlineMeeting": {"joinUrl": getattr(online, "join_url", None)} if online else None,
        "attendees": attendees,
        "bodyPreview": getattr(evt, "body_preview", None),
        "body": _body_to_dict(getattr(evt, "body", None)),
        "showAs": _enum_value(getattr(evt, "show_as", None)),
        "sensitivity": _enum_value(getattr(evt, "sensitivity", None)),
        "recurrence": getattr(evt, "recurrence", None),
        "webLink": getattr(evt, "web_link", None),
    }
    return {**_additional(evt), **base}


def message_rule_predicates_to_dict(pred: Any) -> dict | None:
    if pred is None:
        return None
    base = {
        "bodyContains": list(getattr(pred, "body_contains", None) or []) or None,
        "bodyOrSubjectContains": list(getattr(pred, "body_or_subject_contains", None) or []) or None,
        "categories": list(getattr(pred, "categories", None) or []) or None,
        "fromAddresses": _recipients(getattr(pred, "from_addresses", None)),
        "hasAttachments": getattr(pred, "has_attachments", None),
        "headerContains": list(getattr(pred, "header_contains", None) or []) or None,
        "senderContains": list(getattr(pred, "sender_contains", None) or []) or None,
        "subjectContains": list(getattr(pred, "subject_contains", None) or []) or None,
        "sentToAddresses": _recipients(getattr(pred, "sent_to_addresses", None)),
        "sentToMe": getattr(pred, "sent_to_me", None),
        "sentOnlyToMe": getattr(pred, "sent_only_to_me", None),
        "importance": _enum_value(getattr(pred, "importance", None)),
        "messageActionFlag": _enum_value(getattr(pred, "message_action_flag", None)),
        "sensitivity": _enum_value(getattr(pred, "sensitivity", None)),
    }
    return {**_additional(pred), **base}


def message_rule_actions_to_dict(actions: Any) -> dict | None:
    if actions is None:
        return None
    base = {
        "assignCategories": list(getattr(actions, "assign_categories", None) or []) or None,
        "copyToFolder": getattr(actions, "copy_to_folder", None),
        "delete": getattr(actions, "delete", None),
        "forwardAsAttachmentTo": _recipients(getattr(actions, "forward_as_attachment_to", None)),
        "forwardTo": _recipients(getattr(actions, "forward_to", None)),
        "markAsRead": getattr(actions, "mark_as_read", None),
        "markImportance": _enum_value(getattr(actions, "mark_importance", None)),
        "moveToFolder": getattr(actions, "move_to_folder", None),
        "permanentDelete": getattr(actions, "permanent_delete", None),
        "redirectTo": _recipients(getattr(actions, "redirect_to", None)),
        "stopProcessingRules": getattr(actions, "stop_processing_rules", None),
    }
    return {**_additional(actions), **base}


def message_rule_to_dict(rule: Any) -> dict:
    base = {
        "id": getattr(rule, "id", None),
        "displayName": getattr(rule, "display_name", None),
        "sequence": getattr(rule, "sequence", None),
        "isEnabled": bool(getattr(rule, "is_enabled", False)),
        "hasError": bool(getattr(rule, "has_error", False)),
        "isReadOnly": bool(getattr(rule, "is_read_only", False)),
        "conditions": message_rule_predicates_to_dict(getattr(rule, "conditions", None)),
        "exceptions": message_rule_predicates_to_dict(getattr(rule, "exceptions", None)),
        "actions": message_rule_actions_to_dict(getattr(rule, "actions", None)),
    }
    return {**_additional(rule), **base}


def _identity_to_dict(identity: Any) -> dict | None:
    if identity is None:
        return None
    return {
        "id": getattr(identity, "id", None),
        "displayName": getattr(identity, "display_name", None),
    }


def _identity_set_user(idset: Any) -> dict | None:
    if idset is None:
        return None
    return _identity_to_dict(getattr(idset, "user", None))


def _chat_attachment_to_dict(att: Any) -> dict:
    return {
        "id": getattr(att, "id", None),
        "contentType": getattr(att, "content_type", None),
        "contentUrl": getattr(att, "content_url", None),
        "name": getattr(att, "name", None),
    }


def _chat_mention_to_dict(m: Any) -> dict:
    return {
        "id": getattr(m, "id", None),
        "mentionText": getattr(m, "mention_text", None),
    }


def _chat_reaction_to_dict(r: Any) -> dict:
    return {
        "reactionType": getattr(r, "reaction_type", None),
        "createdDateTime": getattr(r, "created_date_time", None),
        "user": _identity_set_user(getattr(r, "user", None)),
    }


def chat_message_to_dict(msg: Any) -> dict:
    base = {
        "id": getattr(msg, "id", None),
        "messageType": _enum_value(getattr(msg, "message_type", None)),
        "createdDateTime": getattr(msg, "created_date_time", None),
        "lastModifiedDateTime": getattr(msg, "last_modified_date_time", None),
        "deletedDateTime": getattr(msg, "deleted_date_time", None),
        "importance": _enum_value(getattr(msg, "importance", None)),
        "subject": getattr(msg, "subject", None),
        "from": _identity_set_user(getattr(msg, "from_", None)),
        "body": _body_to_dict(getattr(msg, "body", None)),
        "attachments": [_chat_attachment_to_dict(a) for a in (getattr(msg, "attachments", None) or [])],
        "mentions": [_chat_mention_to_dict(m) for m in (getattr(msg, "mentions", None) or [])],
        "reactions": [_chat_reaction_to_dict(r) for r in (getattr(msg, "reactions", None) or [])],
        "webUrl": getattr(msg, "web_url", None),
        "etag": getattr(msg, "etag", None),
    }
    return {**_additional(msg), **base}


def _member_display_names(members: Any) -> list[str]:
    return [
        getattr(m, "display_name", None)
        for m in (members or [])
        if getattr(m, "display_name", None)
    ]


def chat_to_dict(chat: Any) -> dict:
    base = {
        "id": getattr(chat, "id", None),
        "chatType": _enum_value(getattr(chat, "chat_type", None)),
        "topic": getattr(chat, "topic", None),
        "lastUpdatedDateTime": getattr(chat, "last_updated_date_time", None),
        "members": _member_display_names(getattr(chat, "members", None)),
        "webUrl": getattr(chat, "web_url", None),
    }
    return {**_additional(chat), **base}


def team_to_dict(team: Any) -> dict:
    base = {
        "id": getattr(team, "id", None),
        "displayName": getattr(team, "display_name", None),
        "description": getattr(team, "description", None),
    }
    return {**_additional(team), **base}


def channel_to_dict(ch: Any) -> dict:
    base = {
        "id": getattr(ch, "id", None),
        "displayName": getattr(ch, "display_name", None),
        "description": getattr(ch, "description", None),
        "membershipType": _enum_value(getattr(ch, "membership_type", None)),
        "webUrl": getattr(ch, "web_url", None),
    }
    return {**_additional(ch), **base}
