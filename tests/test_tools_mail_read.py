from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from msgraph_mcp.graph.errors import GraphValidationError
from msgraph_mcp.tools import mail_read


def _fake_message(*, id_="m1", subject="Hi"):
    return SimpleNamespace(
        id=id_,
        subject=subject,
        from_=SimpleNamespace(
            email_address=SimpleNamespace(name="A", address="a@x.com")
        ),
        to_recipients=[],
        cc_recipients=[],
        received_date_time="2026-05-19T14:00:00Z",
        sent_date_time="2026-05-19T13:59:00Z",
        body_preview="hi",
        body=SimpleNamespace(content_type=SimpleNamespace(value="text"), content="hi"),
        is_read=False,
        is_draft=False,
        has_attachments=False,
        importance=SimpleNamespace(value="normal"),
        flag=SimpleNamespace(flag_status=SimpleNamespace(value="notFlagged")),
        categories=[],
        conversation_id="c1",
        web_link="https://outlook.example",
        parent_folder_id=None,
        additional_data={},
    )


def _fake_collection(items, next_link=None):
    return SimpleNamespace(value=items, odata_next_link=next_link)


@pytest.mark.asyncio
async def test_list_messages_default_folder_inbox():
    graph = MagicMock()
    mb = MagicMock()
    folder = MagicMock()
    folder.messages.get = AsyncMock(return_value=_fake_collection([_fake_message()]))
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=folder)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.list_messages(graph=graph)

    mb.mail_folders.by_mail_folder_id.assert_called_once_with("inbox")
    assert len(result["items"]) == 1
    assert result["items"][0]["id"] == "m1"
    assert result["next_page_token"] is None


@pytest.mark.asyncio
async def test_list_messages_with_folder_id_and_mailbox():
    graph = MagicMock()
    mb = MagicMock()
    folder = MagicMock()
    folder.messages.get = AsyncMock(
        return_value=_fake_collection(
            [],
            next_link="https://graph.microsoft.com/v1.0/users/x/mailFolders/AAA/messages?$skip=25",
        )
    )
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=folder)
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.list_messages(
        graph=graph, folder_id="AAA", mailbox="alice@example.com", limit=50
    )

    graph.mailbox.assert_called_once_with("alice@example.com")
    mb.mail_folders.by_mail_folder_id.assert_called_once_with("AAA")
    assert result["next_page_token"] is not None


@pytest.mark.asyncio
async def test_list_messages_rejects_bad_limit():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_read.list_messages(graph=graph, limit=200)


def _list_messages_mock_setup():
    graph = MagicMock()
    mb = MagicMock()
    folder = MagicMock()
    folder.messages.get = AsyncMock(return_value=_fake_collection([_fake_message()]))
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=folder)
    graph.mailbox = MagicMock(return_value=mb)
    return graph, folder


def _captured_filter(folder):
    rc = folder.messages.get.await_args.kwargs["request_configuration"]
    return getattr(rc.query_parameters, "filter", None)


@pytest.mark.asyncio
async def test_list_messages_default_passes_no_filter():
    graph, folder = _list_messages_mock_setup()
    await mail_read.list_messages(graph=graph)
    assert _captured_filter(folder) is None


@pytest.mark.asyncio
async def test_list_messages_unread_only_sets_isread_filter():
    graph, folder = _list_messages_mock_setup()
    await mail_read.list_messages(graph=graph, unread_only=True)
    assert _captured_filter(folder) == "isRead eq false"


@pytest.mark.asyncio
async def test_list_messages_raw_filter_passes_through():
    graph, folder = _list_messages_mock_setup()
    await mail_read.list_messages(
        graph=graph, filter="from/emailAddress/address eq 'a@x.com'"
    )
    assert _captured_filter(folder) == "(from/emailAddress/address eq 'a@x.com')"


@pytest.mark.asyncio
async def test_list_messages_combines_unread_and_filter_with_and():
    graph, folder = _list_messages_mock_setup()
    await mail_read.list_messages(
        graph=graph, unread_only=True, filter="contains(subject,'report')"
    )
    assert (
        _captured_filter(folder) == "isRead eq false and (contains(subject,'report'))"
    )


@pytest.mark.asyncio
async def test_search_messages_calls_root_messages_with_search_param():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.get = AsyncMock(
        return_value=_fake_collection([_fake_message(id_="s1")])
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.search_messages(graph=graph, query="from:alice")

    mb.messages.get.assert_awaited_once()
    # Verify $search was set on the request configuration
    await_args = mb.messages.get.await_args
    assert await_args is not None
    rc = await_args.kwargs.get("request_configuration")
    assert rc is not None
    # Search value should appear in the query parameters
    qp = getattr(rc, "query_parameters", None)
    assert qp is not None
    assert getattr(qp, "search", None) == "from:alice"
    assert result["items"][0]["id"] == "s1"


@pytest.mark.asyncio
async def test_get_message_returns_trimmed_message_with_body():
    graph = MagicMock()
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(get=AsyncMock(return_value=_fake_message(id_="m99")))
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.get_message(
        graph=graph, message_id="m99", include_body=True
    )

    mb.messages.by_message_id.assert_called_once_with("m99")
    assert result["id"] == "m99"
    assert "body" in result
    assert result["body"] == "hi"


@pytest.mark.asyncio
async def test_list_attachments_returns_trimmed_list():
    graph = MagicMock()
    mb = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="x.pdf",
        content_type="application/pdf",
        size=10,
        is_inline=False,
        additional_data={},
    )
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(
            attachments=MagicMock(get=AsyncMock(return_value=_fake_collection([att])))
        )
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.list_attachments(graph=graph, message_id="m1")

    assert result["items"][0]["name"] == "x.pdf"


@pytest.mark.asyncio
async def test_download_attachment_returns_base64_content():
    graph = MagicMock()
    mb = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="x.pdf",
        content_type="application/pdf",
        size=4,
        is_inline=False,
        content_bytes=b"\x00\x01\x02\x03",
        additional_data={},
    )
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(
            attachments=MagicMock(
                by_attachment_id=MagicMock(
                    return_value=MagicMock(get=AsyncMock(return_value=att))
                )
            )
        )
    )
    graph.mailbox = MagicMock(return_value=mb)

    result = await mail_read.download_attachment(
        graph=graph, message_id="m1", attachment_id="a1"
    )

    assert isinstance(result, dict)
    assert result["name"] == "x.pdf"
    assert result["content_base64"] is not None


PNG = b"\x89PNG\r\n\x1a\n\x00\x00\x00\x0dIHDR"


def _wire_attachment(graph, att):
    mb = MagicMock()
    mb.messages.by_message_id = MagicMock(
        return_value=MagicMock(
            attachments=MagicMock(
                by_attachment_id=MagicMock(
                    return_value=MagicMock(get=AsyncMock(return_value=att))
                )
            )
        )
    )
    graph.mailbox = MagicMock(return_value=mb)


@pytest.mark.asyncio
async def test_download_attachment_image_returns_image_content():
    import base64 as b64

    from mcp.server.fastmcp.utilities.types import Image

    graph = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="shot.png",
        content_type="image/png",
        size=len(PNG),
        is_inline=False,
        content_bytes=PNG,
        additional_data={},
    )
    _wire_attachment(graph, att)

    result = await mail_read.download_attachment(
        graph=graph, message_id="m1", attachment_id="a1"
    )

    assert isinstance(result, list) and len(result) == 2
    meta, image = result
    assert meta["name"] == "shot.png"
    assert meta["content_type"] == "image/png"
    assert isinstance(image, Image)
    block = image.to_image_content()
    assert block.mimeType == "image/png"
    assert b64.b64decode(block.data) == PNG


@pytest.mark.asyncio
async def test_download_attachment_save_path_writes_file(tmp_path):
    graph = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="x.pdf",
        content_type="application/pdf",
        size=4,
        is_inline=False,
        content_bytes=b"\x00\x01\x02\x03",
        additional_data={},
    )
    _wire_attachment(graph, att)

    target = tmp_path / "out.pdf"
    result = await mail_read.download_attachment(
        graph=graph, message_id="m1", attachment_id="a1", save_path=str(target)
    )

    assert isinstance(result, dict)
    assert result["path"] == str(target)
    assert result["name"] == "x.pdf"
    assert "content_base64" not in result
    assert target.read_bytes() == b"\x00\x01\x02\x03"


@pytest.mark.asyncio
async def test_download_attachment_save_path_directory_uses_attachment_name(tmp_path):
    graph = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="report.pdf",
        content_type="application/pdf",
        size=4,
        is_inline=False,
        content_bytes=b"\x00\x01\x02\x03",
        additional_data={},
    )
    _wire_attachment(graph, att)

    result = await mail_read.download_attachment(
        graph=graph, message_id="m1", attachment_id="a1", save_path=str(tmp_path)
    )

    assert isinstance(result, dict)
    assert result["path"] == str(tmp_path / "report.pdf")
    assert (tmp_path / "report.pdf").read_bytes() == b"\x00\x01\x02\x03"


@pytest.mark.asyncio
async def test_download_attachment_save_path_decodes_str_content_bytes(tmp_path):
    import base64 as b64

    graph = MagicMock()
    att = SimpleNamespace(
        id="a1",
        name="x.bin",
        content_type="application/octet-stream",
        size=4,
        is_inline=False,
        content_bytes=b64.b64encode(b"\x00\x01\x02\x03").decode("ascii"),
        additional_data={},
    )
    _wire_attachment(graph, att)

    result = await mail_read.download_attachment(
        graph=graph,
        message_id="m1",
        attachment_id="a1",
        save_path=str(tmp_path / "x.bin"),
    )

    assert (tmp_path / "x.bin").read_bytes() == b"\x00\x01\x02\x03"
    assert isinstance(result, dict)
    assert result["size_bytes"] == 4
