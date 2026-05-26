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


@pytest.mark.asyncio
async def test_create_folder_top_level():
    graph = MagicMock()
    mb = MagicMock()
    created = _fake_folder(id_="new-id", name="Test notifications")
    mb.mail_folders.post = AsyncMock(return_value=created)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_folders.create_folder(
        graph=graph, display_name="Test notifications"
    )
    assert result["display_name"] == "Test notifications"
    assert result["id"] == "new-id"
    # Verify the right endpoint was hit (top-level)
    assert mb.mail_folders.post.await_count == 1
    posted_body = mb.mail_folders.post.await_args.args[0]
    assert posted_body.display_name == "Test notifications"


@pytest.mark.asyncio
async def test_create_folder_child():
    graph = MagicMock()
    mb = MagicMock()
    created = _fake_folder(id_="child-id", name="Sub")
    child_endpoint = MagicMock(child_folders=MagicMock(post=AsyncMock(return_value=created)))
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=child_endpoint)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_folders.create_folder(
        graph=graph, display_name="Sub", parent_folder_id="parent-id"
    )
    assert result["id"] == "child-id"
    mb.mail_folders.by_mail_folder_id.assert_called_once_with("parent-id")
    child_endpoint.child_folders.post.assert_awaited_once()
    posted_body = child_endpoint.child_folders.post.await_args.args[0]
    assert posted_body.display_name == "Sub"


@pytest.mark.asyncio
async def test_create_folder_validates_empty_name():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_folders.create_folder(graph=graph, display_name="")


@pytest.mark.asyncio
async def test_create_folder_validates_whitespace_name():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_folders.create_folder(graph=graph, display_name="   ")


def _empty_folder(id_="f1", name="Empty"):
    return SimpleNamespace(
        id=id_,
        display_name=name,
        parent_folder_id=None,
        total_item_count=0,
        unread_item_count=0,
        child_folder_count=0,
        additional_data={},
    )


def _build_folder_endpoint(folder):
    """Wire graph.mailbox(x).mail_folders.by_mail_folder_id(id) to a MagicMock
    whose .get() returns `folder` and .delete() is awaitable.
    """
    graph = MagicMock()
    mb = MagicMock()
    by_id = MagicMock()
    by_id.get = AsyncMock(return_value=folder)
    by_id.delete = AsyncMock(return_value=None)
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=by_id)
    graph.mailbox = MagicMock(return_value=mb)
    return graph, mb, by_id


@pytest.mark.asyncio
async def test_delete_folder_empty_calls_delete():
    graph, mb, by_id = _build_folder_endpoint(_empty_folder(id_="fid"))
    result = await mail_folders.delete_folder(graph=graph, folder_id="fid")
    assert result == {"deleted": True, "id": "fid"}
    mb.mail_folders.by_mail_folder_id.assert_called_with("fid")
    by_id.get.assert_awaited_once()
    by_id.delete.assert_awaited_once()


@pytest.mark.asyncio
async def test_delete_folder_refuses_when_messages_present():
    from outlook_mcp.graph.errors import GraphValidationError

    folder = _empty_folder(id_="fid")
    folder.total_item_count = 3
    graph, mb, by_id = _build_folder_endpoint(folder)
    with pytest.raises(GraphValidationError, match="non-empty folder"):
        await mail_folders.delete_folder(graph=graph, folder_id="fid")
    by_id.delete.assert_not_awaited()


@pytest.mark.asyncio
async def test_delete_folder_refuses_when_child_folders_present():
    from outlook_mcp.graph.errors import GraphValidationError

    folder = _empty_folder(id_="fid")
    folder.child_folder_count = 1
    graph, mb, by_id = _build_folder_endpoint(folder)
    with pytest.raises(GraphValidationError, match="non-empty folder"):
        await mail_folders.delete_folder(graph=graph, folder_id="fid")
    by_id.delete.assert_not_awaited()


@pytest.mark.asyncio
async def test_delete_folder_requires_folder_id():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_folders.delete_folder(graph=graph, folder_id="")
