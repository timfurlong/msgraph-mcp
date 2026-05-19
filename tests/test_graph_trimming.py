from tests.fixtures import load

from outlook_mcp.graph import trimming


def test_trim_message_minimal():
    raw = load("message_minimal.json")
    t = trimming.trim_message(raw, include_body=False, include_raw=False)
    assert t["id"] == "AAMkAGI-msg-id"
    assert t["subject"] == "Hello"
    assert t["from"] == {"name": "Alice", "address": "alice@example.com"}
    assert t["to"] == [{"name": "Bob", "address": "bob@example.com"}]
    assert t["cc"] == []
    assert t["received"] == "2026-05-19T14:00:00Z"
    assert t["snippet"] == "Hello there"
    assert "body" not in t  # excluded when include_body=False
    assert t["is_read"] is False
    assert t["flag"] == "notFlagged"
    assert t["categories"] == []
    assert "raw" not in t


def test_trim_message_with_body_and_raw():
    raw = load("message_minimal.json")
    t = trimming.trim_message(raw, include_body=True, include_raw=True)
    assert t["body"] == "Hello there\nFull body."
    assert t["body_type"] == "text"
    assert t["raw"] == raw


def test_trim_message_handles_missing_optional_fields():
    minimal = {
        "id": "x",
        "subject": "y",
        "isRead": True,
        "isDraft": False,
        "hasAttachments": False,
    }
    t = trimming.trim_message(minimal, include_body=False, include_raw=False)
    assert t["from"] is None
    assert t["to"] == []
    assert t["cc"] == []
    assert t["received"] is None
    assert t["flag"] == "notFlagged"


def test_trim_event_minimal():
    raw = load("event_minimal.json")
    t = trimming.trim_event(raw, include_body=False, include_raw=False)
    assert t["subject"] == "Standup"
    assert t["organizer"] == {"name": "Alice", "address": "alice@example.com"}
    assert t["start"] == {"date_time": "2026-05-19T15:00:00", "time_zone": "America/Los_Angeles"}
    assert t["location"] == "Zoom"
    assert t["is_online_meeting"] is True
    assert t["online_meeting_url"] == "https://zoom.us/j/123"
    assert t["attendees"][0]["response"] == "accepted"
    assert "body" not in t


def test_trim_event_with_body_and_raw():
    raw = load("event_minimal.json")
    t = trimming.trim_event(raw, include_body=True, include_raw=True)
    assert t["body"] == "<p>Daily standup</p>"
    assert t["raw"] == raw


def test_trim_folder():
    raw = {
        "id": "f1",
        "displayName": "Inbox",
        "parentFolderId": "p1",
        "totalItemCount": 42,
        "unreadItemCount": 3,
        "childFolderCount": 0,
    }
    t = trimming.trim_folder(raw, include_raw=False)
    assert t == {
        "id": "f1",
        "display_name": "Inbox",
        "parent_folder_id": "p1",
        "total_item_count": 42,
        "unread_item_count": 3,
        "child_folder_count": 0,
    }


def test_trim_attachment_list_view():
    raw = {
        "id": "a1",
        "name": "doc.pdf",
        "contentType": "application/pdf",
        "size": 12345,
        "isInline": False,
    }
    t = trimming.trim_attachment_list(raw, include_raw=False)
    assert t == {
        "id": "a1",
        "name": "doc.pdf",
        "content_type": "application/pdf",
        "size_bytes": 12345,
        "is_inline": False,
    }


def test_trim_user():
    raw = {
        "id": "u1",
        "displayName": "Alice",
        "userPrincipalName": "alice@example.com",
        "mail": "alice@example.com",
        "jobTitle": "Engineer",
    }
    t = trimming.trim_user(raw, include_raw=False)
    assert t["display_name"] == "Alice"
    assert t["user_principal_name"] == "alice@example.com"
    assert t["mail"] == "alice@example.com"
    assert t["job_title"] == "Engineer"
