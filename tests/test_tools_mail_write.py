import base64
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import mail_write


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


@pytest.mark.asyncio
async def test_send_message_minimal():
    graph = MagicMock()
    mb = MagicMock()
    mb.send_mail.post = AsyncMock(return_value=None)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.send_message(
        graph=graph, to=["bob@example.com"], subject="Hi", body="Hello"
    )

    mb.send_mail.post.assert_awaited_once()
    assert result == {"status": "sent"}


@pytest.mark.asyncio
async def test_send_message_rejects_oversize_attachment():
    graph = MagicMock()
    huge = _b64(b"X" * (3 * 1024 * 1024 + 1))  # > 3 MB raw
    with pytest.raises(GraphValidationError):
        await mail_write.send_message(
            graph=graph,
            to=["bob@example.com"],
            subject="Hi",
            body="Hello",
            attachments=[
                {"name": "big.bin", "content_b64": huge, "content_type": "application/octet-stream"}
            ],
        )


@pytest.mark.asyncio
async def test_create_draft_returns_id():
    graph = MagicMock()
    mb = MagicMock()
    draft = SimpleNamespace(
        id="d1", subject="Draft", from_=None, to_recipients=[], cc_recipients=[],
        received_date_time=None, sent_date_time=None, body_preview="", body=None,
        is_read=True, is_draft=True, has_attachments=False, importance=None, flag=None,
        categories=[], conversation_id=None, web_link=None, parent_folder_id=None,
        additional_data={},
    )
    mb.messages.post = AsyncMock(return_value=draft)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.create_draft(
        graph=graph, to=["bob@example.com"], subject="Draft", body="Body"
    )
    assert result["id"] == "d1"
    assert result["is_draft"] is True


@pytest.mark.asyncio
async def test_reply_message_calls_reply_endpoint():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(reply=MagicMock(post=AsyncMock(return_value=None)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.reply_message(graph=graph, message_id="m1", comment="thanks")

    mb.messages.by_message_id.assert_called_once_with("m1")
    assert result == {"status": "sent"}


@pytest.mark.asyncio
async def test_reply_all_message_calls_reply_all_endpoint():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(reply_all=MagicMock(post=AsyncMock(return_value=None)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.reply_all_message(graph=graph, message_id="m1", comment="ack")
    assert result == {"status": "sent"}


@pytest.mark.asyncio
async def test_forward_message_calls_forward_endpoint():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(forward=MagicMock(post=AsyncMock(return_value=None)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.forward_message(
        graph=graph, message_id="m1", to=["carol@example.com"], comment="fyi"
    )
    assert result == {"status": "sent"}


@pytest.mark.asyncio
async def test_update_message_sets_is_read():
    graph = MagicMock()
    mb = MagicMock()
    updated = SimpleNamespace(
        id="m1", subject="x", from_=None, to_recipients=[], cc_recipients=[],
        received_date_time=None, sent_date_time=None, body_preview="", body=None,
        is_read=True, is_draft=False, has_attachments=False, importance=None, flag=None,
        categories=[], conversation_id=None, web_link=None, parent_folder_id=None,
        additional_data={},
    )
    mb.messages.by_message_id = MagicMock(return_value=MagicMock(patch=AsyncMock(return_value=updated)))
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.update_message(graph=graph, message_id="m1", is_read=True)
    assert result["is_read"] is True


@pytest.mark.asyncio
async def test_delete_message_returns_status():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(return_value=MagicMock(delete=AsyncMock(return_value=None)))
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_write.delete_message(graph=graph, message_id="m1")
    assert result == {"status": "deleted"}
