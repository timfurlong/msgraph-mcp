from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.tools import mail_actions


@pytest.mark.asyncio
async def test_archive_message_delegates_to_move():
    graph = MagicMock()
    with patch(
        "outlook_mcp.tools.mail_actions.move_message",
        new=AsyncMock(return_value={"id": "m1"}),
    ) as p:
        result = await mail_actions.archive_message(graph=graph, message_id="m1")
    p.assert_awaited_once_with(
        graph=graph, message_id="m1", destination="archive", mailbox=None, include_raw=False
    )
    assert result == {"id": "m1"}


@pytest.mark.asyncio
async def test_mark_read_delegates_to_update_message():
    graph = MagicMock()
    with patch(
        "outlook_mcp.tools.mail_actions.update_message",
        new=AsyncMock(return_value={"is_read": True}),
    ) as p:
        result = await mail_actions.mark_read(graph=graph, message_id="m1")
    p.assert_awaited_once_with(
        graph=graph, message_id="m1", is_read=True, mailbox=None, include_raw=False
    )
    assert result["is_read"] is True


@pytest.mark.asyncio
async def test_mark_unread_delegates_to_update_message():
    graph = MagicMock()
    with patch(
        "outlook_mcp.tools.mail_actions.update_message",
        new=AsyncMock(return_value={"is_read": False}),
    ) as p:
        await mail_actions.mark_unread(graph=graph, message_id="m1")
    p.assert_awaited_once_with(
        graph=graph, message_id="m1", is_read=False, mailbox=None, include_raw=False
    )


@pytest.mark.asyncio
async def test_flag_message_delegates_to_update_message():
    graph = MagicMock()
    with patch(
        "outlook_mcp.tools.mail_actions.update_message",
        new=AsyncMock(return_value={"flag": "flagged"}),
    ) as p:
        await mail_actions.flag_message(graph=graph, message_id="m1")
    p.assert_awaited_once_with(
        graph=graph, message_id="m1", flag="flagged", mailbox=None, include_raw=False
    )


@pytest.mark.asyncio
async def test_unflag_message_delegates_to_update_message():
    graph = MagicMock()
    with patch(
        "outlook_mcp.tools.mail_actions.update_message",
        new=AsyncMock(return_value={"flag": "notFlagged"}),
    ) as p:
        await mail_actions.unflag_message(graph=graph, message_id="m1")
    p.assert_awaited_once_with(
        graph=graph, message_id="m1", flag="notFlagged", mailbox=None, include_raw=False
    )
