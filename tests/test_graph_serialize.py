from types import SimpleNamespace

from outlook_mcp.graph import serialize
from outlook_mcp.graph.serialize import (
    chat_message_to_dict,
    chat_to_dict,
    team_to_dict,
    channel_to_dict,
)


def test_user_to_dict():
    user = SimpleNamespace(
        id="u1",
        display_name="Alice",
        user_principal_name="alice@example.com",
        mail="alice@example.com",
        job_title="Engineer",
        additional_data={"extra": "x"},
    )
    d = serialize.user_to_dict(user)
    assert d["id"] == "u1"
    assert d["displayName"] == "Alice"
    assert d["userPrincipalName"] == "alice@example.com"
    assert d["extra"] == "x"


def test_recipient_to_dict():
    addr = SimpleNamespace(name="Bob", address="bob@example.com")
    rec = SimpleNamespace(email_address=addr)
    assert serialize.recipient_to_dict(rec) == {
        "emailAddress": {"name": "Bob", "address": "bob@example.com"}
    }


def test_message_to_dict_minimal():
    body = SimpleNamespace(content_type=SimpleNamespace(value="text"), content="Hi")
    msg = SimpleNamespace(
        id="m1",
        subject="Hello",
        from_=SimpleNamespace(email_address=SimpleNamespace(name="A", address="a@x.com")),
        to_recipients=[SimpleNamespace(email_address=SimpleNamespace(name="B", address="b@x.com"))],
        cc_recipients=[],
        received_date_time="2026-05-19T14:00:00Z",
        sent_date_time="2026-05-19T13:59:00Z",
        body_preview="Hi there",
        body=body,
        is_read=False,
        is_draft=False,
        has_attachments=False,
        importance=SimpleNamespace(value="normal"),
        flag=SimpleNamespace(flag_status=SimpleNamespace(value="notFlagged")),
        categories=[],
        conversation_id="c1",
        web_link="https://outlook.example",
        additional_data={},
    )
    d = serialize.message_to_dict(msg)
    assert d["id"] == "m1"
    assert d["subject"] == "Hello"
    assert d["from"] == {"emailAddress": {"name": "A", "address": "a@x.com"}}
    assert d["toRecipients"] == [{"emailAddress": {"name": "B", "address": "b@x.com"}}]
    assert d["body"] == {"contentType": "text", "content": "Hi"}
    assert d["flag"] == {"flagStatus": "notFlagged"}


def test_message_to_dict_surfaces_parent_folder_and_inference():
    msg = SimpleNamespace(
        id="m1",
        parent_folder_id="AAMkAGI-folder-id",
        inference_classification=SimpleNamespace(value="focused"),
        additional_data={},
    )
    d = serialize.message_to_dict(msg)
    assert d["parentFolderId"] == "AAMkAGI-folder-id"
    assert d["inferenceClassification"] == "focused"


def test_folder_to_dict():
    f = SimpleNamespace(
        id="f1",
        display_name="Inbox",
        parent_folder_id="p1",
        total_item_count=42,
        unread_item_count=3,
        child_folder_count=0,
        additional_data={},
    )
    d = serialize.folder_to_dict(f)
    assert d["displayName"] == "Inbox"
    assert d["totalItemCount"] == 42


def test_attachment_to_dict():
    a = SimpleNamespace(
        id="a1",
        name="doc.pdf",
        content_type="application/pdf",
        size=12345,
        is_inline=False,
        additional_data={},
    )
    d = serialize.attachment_to_dict(a)
    assert d["name"] == "doc.pdf"
    assert d["size"] == 12345


def test_attachment_to_dict_with_content_bytes():
    a = SimpleNamespace(
        id="a1",
        name="doc.pdf",
        content_type="application/pdf",
        size=12345,
        is_inline=False,
        content_bytes=b"raw-bytes",
        additional_data={},
    )
    d = serialize.attachment_to_dict(a, include_content=True)
    # content_bytes are base64-encoded on the wire; we store them as a str
    assert d["contentBytes"] is not None


def test_message_rule_predicates_to_dict_minimal():
    pred = SimpleNamespace(
        body_contains=None,
        body_or_subject_contains=None,
        categories=None,
        from_addresses=None,
        has_attachments=None,
        header_contains=None,
        sender_contains=["example.com"],
        subject_contains=["Security Scan"],
        sent_to_addresses=None,
        sent_to_me=None,
        sent_only_to_me=None,
        importance=None,
        message_action_flag=None,
        sensitivity=None,
        within_size_range=None,
        additional_data={},
    )
    d = serialize.message_rule_predicates_to_dict(pred)
    assert d["senderContains"] == ["example.com"]
    assert d["subjectContains"] == ["Security Scan"]
    assert d["bodyContains"] is None
    assert d["fromAddresses"] == []


def test_message_rule_actions_to_dict_minimal():
    actions = SimpleNamespace(
        assign_categories=None,
        copy_to_folder=None,
        delete=None,
        forward_as_attachment_to=None,
        forward_to=None,
        mark_as_read=True,
        mark_importance=None,
        move_to_folder="AAMkFolderId",
        permanent_delete=None,
        redirect_to=None,
        stop_processing_rules=True,
        additional_data={},
    )
    d = serialize.message_rule_actions_to_dict(actions)
    assert d["moveToFolder"] == "AAMkFolderId"
    assert d["markAsRead"] is True
    assert d["stopProcessingRules"] is True
    assert d["delete"] is None
    assert d["forwardTo"] == []


def test_message_rule_to_dict_minimal():
    rule = SimpleNamespace(
        id="rule-1",
        display_name="Test notifications",
        sequence=1,
        is_enabled=True,
        has_error=False,
        is_read_only=False,
        conditions=None,
        exceptions=None,
        actions=None,
        additional_data={},
    )
    d = serialize.message_rule_to_dict(rule)
    assert d["id"] == "rule-1"
    assert d["displayName"] == "Test notifications"
    assert d["sequence"] == 1
    assert d["isEnabled"] is True
    assert d["conditions"] is None
    assert d["actions"] is None


def _fake_chat_message():
    return SimpleNamespace(
        id="1700000000000",
        message_type=SimpleNamespace(value="message"),
        created_date_time="2026-07-01T10:00:00Z",
        last_modified_date_time="2026-07-01T10:00:00Z",
        deleted_date_time=None,
        importance=SimpleNamespace(value="normal"),
        subject=None,
        from_=SimpleNamespace(user=SimpleNamespace(id="u1", display_name="Alice")),
        body=SimpleNamespace(content_type=SimpleNamespace(value="html"), content="<p>hi</p>"),
        attachments=[SimpleNamespace(id="a1", content_type="reference", content_url="https://example.com/f.docx", name="f.docx")],
        mentions=[SimpleNamespace(id=0, mention_text="Bob")],
        reactions=[SimpleNamespace(reaction_type="like", created_date_time="2026-07-01T10:05:00Z", user=SimpleNamespace(user=SimpleNamespace(id="u2", display_name="Bob")))],
        web_url="https://teams.example/msg/1",
        etag="1700000000000",
        additional_data={},
    )


def test_chat_message_to_dict_core_fields():
    d = chat_message_to_dict(_fake_chat_message())
    assert d["id"] == "1700000000000"
    assert d["messageType"] == "message"
    assert d["from"] == {"id": "u1", "displayName": "Alice"}
    assert d["body"] == {"contentType": "html", "content": "<p>hi</p>"}
    assert d["attachments"][0]["name"] == "f.docx"
    assert d["mentions"][0]["mentionText"] == "Bob"
    assert d["reactions"][0]["reactionType"] == "like"


def test_chat_message_to_dict_handles_missing_from_and_body():
    msg = SimpleNamespace(id="x", message_type=None, created_date_time=None,
                          last_modified_date_time=None, deleted_date_time=None,
                          importance=None, subject=None, from_=None, body=None,
                          attachments=None, mentions=None, reactions=None,
                          web_url=None, etag=None, additional_data={})
    d = chat_message_to_dict(msg)
    assert d["from"] is None
    assert d["body"] is None
    assert d["attachments"] == []


def test_chat_to_dict_flattens_member_names():
    chat = SimpleNamespace(
        id="c1", chat_type=SimpleNamespace(value="group"), topic="Launch",
        last_updated_date_time="2026-07-01T09:00:00Z",
        last_message_preview=SimpleNamespace(created_date_time="2026-07-03T17:46:28Z"),
        members=[SimpleNamespace(display_name="Alice"), SimpleNamespace(display_name="Bob"), SimpleNamespace(display_name=None)],
        web_url="https://teams.example/chat/c1", additional_data={},
    )
    d = chat_to_dict(chat)
    assert d["chatType"] == "group"
    assert d["members"] == ["Alice", "Bob"]
    assert d["lastMessagePreview"] == {"createdDateTime": "2026-07-03T17:46:28Z"}


def test_chat_to_dict_missing_last_message_preview_is_none():
    chat = SimpleNamespace(
        id="c1", chat_type=SimpleNamespace(value="oneOnOne"), topic=None,
        last_updated_date_time="2025-12-22T17:46:20Z", last_message_preview=None,
        members=[SimpleNamespace(display_name="Alice")],
        web_url="u", additional_data={},
    )
    assert chat_to_dict(chat)["lastMessagePreview"] is None


def test_team_to_dict():
    team = SimpleNamespace(id="t1", display_name="Eng", description="Engineering", additional_data={})
    assert team_to_dict(team) == {"id": "t1", "displayName": "Eng", "description": "Engineering"}


def test_channel_to_dict():
    ch = SimpleNamespace(id="ch1", display_name="General", description=None,
                         membership_type=SimpleNamespace(value="standard"),
                         web_url="https://teams.example/ch1", additional_data={})
    d = channel_to_dict(ch)
    assert d["displayName"] == "General"
    assert d["membershipType"] == "standard"
