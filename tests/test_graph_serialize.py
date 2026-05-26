from types import SimpleNamespace

from outlook_mcp.graph import serialize


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
