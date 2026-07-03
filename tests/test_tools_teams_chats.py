from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import teams_chats


def _fake_collection(items, next_link=None):
    return SimpleNamespace(value=items, odata_next_link=next_link)


def _fake_chat(id_="c1"):
    return SimpleNamespace(
        id=id_, chat_type=SimpleNamespace(value="group"), topic="Launch",
        last_updated_date_time="2026-07-01T09:00:00Z",
        members=[SimpleNamespace(display_name="Alice")],
        web_url="https://teams.example/chat/c1", additional_data={},
    )


def _fake_chat_message(id_="m1"):
    return SimpleNamespace(
        id=id_, message_type=SimpleNamespace(value="message"),
        created_date_time="2026-07-01T10:00:00Z", last_modified_date_time="2026-07-01T10:00:00Z",
        deleted_date_time=None, importance=SimpleNamespace(value="normal"), subject=None,
        from_=SimpleNamespace(user=SimpleNamespace(id="u1", display_name="Alice")),
        body=SimpleNamespace(content_type=SimpleNamespace(value="text"), content="hi"),
        attachments=[], mentions=[], reactions=[], web_url="https://teams.example/msg/1",
        etag="1", additional_data={},
    )


@pytest.mark.asyncio
async def test_list_chats_returns_trimmed_and_expands_members():
    graph = MagicMock()
    graph.raw.me.chats.get = AsyncMock(return_value=_fake_collection([_fake_chat()]))

    result = await teams_chats.list_chats(graph=graph)

    graph.raw.me.chats.get.assert_awaited_once()
    rc = graph.raw.me.chats.get.await_args.kwargs["request_configuration"]
    assert "members" in (getattr(rc.query_parameters, "expand", None) or [])
    assert result["items"][0]["id"] == "c1"
    assert result["items"][0]["members"] == ["Alice"]
    assert result["next_page_token"] is None


@pytest.mark.asyncio
async def test_list_chats_rejects_bad_limit():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_chats.list_chats(graph=graph, limit=200)


@pytest.mark.asyncio
async def test_list_chats_paginates_with_page_token():
    graph = MagicMock()
    with_url = MagicMock()
    with_url.get = AsyncMock(return_value=_fake_collection([_fake_chat()], next_link=None))
    graph.raw.me.chats.with_url = MagicMock(return_value=with_url)
    # A valid token is base64url of a graph.microsoft.com URL.
    from outlook_mcp.graph.pagination import encode_next_link
    token = encode_next_link("https://graph.microsoft.com/v1.0/me/chats?$skip=25")

    result = await teams_chats.list_chats(graph=graph, page_token=token)

    graph.raw.me.chats.with_url.assert_called_once()
    assert result["items"][0]["id"] == "c1"


@pytest.mark.asyncio
async def test_list_chat_messages_routes_by_chat_id_and_trims():
    graph = MagicMock()
    messages = MagicMock()
    messages.get = AsyncMock(return_value=_fake_collection([_fake_chat_message()]))
    graph.raw.chats.by_chat_id = MagicMock(return_value=MagicMock(messages=messages))

    result = await teams_chats.list_chat_messages(graph=graph, chat_id="c1")

    graph.raw.chats.by_chat_id.assert_called_once_with("c1")
    assert result["items"][0]["id"] == "m1"
    assert "body" not in result["items"][0]


@pytest.mark.asyncio
async def test_list_chat_messages_include_body():
    graph = MagicMock()
    messages = MagicMock()
    messages.get = AsyncMock(return_value=_fake_collection([_fake_chat_message()]))
    graph.raw.chats.by_chat_id = MagicMock(return_value=MagicMock(messages=messages))

    result = await teams_chats.list_chat_messages(graph=graph, chat_id="c1", include_body=True)

    assert result["items"][0]["body"] == "hi"
