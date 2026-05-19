from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools import mail_folders


def _fake_folder(id_="f1", name="Inbox"):
    return SimpleNamespace(
        id=id_, display_name=name, parent_folder_id=None,
        total_item_count=10, unread_item_count=2, child_folder_count=0,
        additional_data={},
    )


@pytest.mark.asyncio
async def test_list_folders_returns_trimmed_list():
    graph = MagicMock()
    mb = MagicMock()
    mb.mail_folders.get = AsyncMock(
        return_value=SimpleNamespace(value=[_fake_folder()], odata_next_link=None)
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_folders.list_folders(graph=graph)
    assert result["items"][0]["display_name"] == "Inbox"


@pytest.mark.asyncio
async def test_move_message_calls_move_endpoint():
    graph = MagicMock()
    mb = MagicMock()
    moved = SimpleNamespace(
        id="m1", subject="x", from_=None, to_recipients=[], cc_recipients=[],
        received_date_time=None, sent_date_time=None, body_preview="", body=None,
        is_read=False, is_draft=False, has_attachments=False, importance=None, flag=None,
        categories=[], conversation_id=None, web_link=None, parent_folder_id="archive",
        additional_data={},
    )
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(move=MagicMock(post=AsyncMock(return_value=moved)))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_folders.move_message(graph=graph, message_id="m1", destination="archive")
    assert result["id"] == "m1"


@pytest.mark.asyncio
async def test_move_message_accepts_well_known_destinations():
    # Just verify validation doesn't reject these.
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(move=MagicMock(post=AsyncMock(return_value=_fake_folder())))
    )
    graph.mailbox = MagicMock(return_value=mb)

    for name in ("archive", "inbox", "junkemail", "deleteditems", "sentitems", "drafts"):
        await mail_folders.move_message(graph=graph, message_id="m1", destination=name)
