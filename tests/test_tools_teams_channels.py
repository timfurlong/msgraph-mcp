from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import teams_channels


def _fake_collection(items, next_link=None):
    return SimpleNamespace(value=items, odata_next_link=next_link)


def _fake_team(id_="t1"):
    return SimpleNamespace(id=id_, display_name="Eng", description="Engineering", additional_data={})


def _fake_channel(id_="ch1"):
    return SimpleNamespace(id=id_, display_name="General", description=None,
                           membership_type=SimpleNamespace(value="standard"),
                           web_url="https://teams.example/ch1", additional_data={})


def _fake_msg(id_="m1"):
    return SimpleNamespace(
        id=id_, message_type=SimpleNamespace(value="message"),
        created_date_time="2026-07-01T10:00:00Z", last_modified_date_time="2026-07-01T10:00:00Z",
        deleted_date_time=None, importance=SimpleNamespace(value="normal"), subject="Post",
        from_=SimpleNamespace(user=SimpleNamespace(id="u1", display_name="Alice")),
        body=SimpleNamespace(content_type=SimpleNamespace(value="text"), content="hi"),
        attachments=[], mentions=[], reactions=[], web_url="https://teams.example/msg/1",
        etag="1", additional_data={},
    )


@pytest.mark.asyncio
async def test_list_joined_teams():
    graph = MagicMock()
    graph.raw.me.joined_teams.get = AsyncMock(return_value=_fake_collection([_fake_team()]))
    result = await teams_channels.list_joined_teams(graph=graph)
    assert result["items"][0]["display_name"] == "Eng"


@pytest.mark.asyncio
async def test_list_channels_routes_by_team_id():
    graph = MagicMock()
    channels = MagicMock()
    channels.get = AsyncMock(return_value=_fake_collection([_fake_channel()]))
    graph.raw.teams.by_team_id = MagicMock(return_value=MagicMock(channels=channels))
    result = await teams_channels.list_channels(graph=graph, team_id="t1")
    graph.raw.teams.by_team_id.assert_called_once_with("t1")
    assert result["items"][0]["membership_type"] == "standard"


@pytest.mark.asyncio
async def test_list_channel_messages_routes_and_trims():
    graph = MagicMock()
    messages = MagicMock()
    messages.get = AsyncMock(return_value=_fake_collection([_fake_msg()]))
    channel = MagicMock(messages=messages)
    team = MagicMock()
    team.channels.by_channel_id = MagicMock(return_value=channel)
    graph.raw.teams.by_team_id = MagicMock(return_value=team)

    result = await teams_channels.list_channel_messages(graph=graph, team_id="t1", channel_id="ch1")

    graph.raw.teams.by_team_id.assert_called_once_with("t1")
    team.channels.by_channel_id.assert_called_once_with("ch1")
    assert result["items"][0]["id"] == "m1"


@pytest.mark.asyncio
async def test_list_channel_messages_caps_limit_at_50():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_channels.list_channel_messages(graph=graph, team_id="t1", channel_id="ch1", limit=51)


@pytest.mark.asyncio
async def test_list_message_replies_routes_by_message_id():
    graph = MagicMock()
    replies = MagicMock()
    replies.get = AsyncMock(return_value=_fake_collection([_fake_msg(id_="r1")]))
    message = MagicMock(replies=replies)
    channel = MagicMock()
    channel.messages.by_chat_message_id = MagicMock(return_value=message)
    team = MagicMock()
    team.channels.by_channel_id = MagicMock(return_value=channel)
    graph.raw.teams.by_team_id = MagicMock(return_value=team)

    result = await teams_channels.list_message_replies(
        graph=graph, team_id="t1", channel_id="ch1", message_id="m1"
    )

    channel.messages.by_chat_message_id.assert_called_once_with("m1")
    assert result["items"][0]["id"] == "r1"


@pytest.mark.asyncio
async def test_list_message_replies_caps_limit_at_50():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_channels.list_message_replies(
            graph=graph, team_id="t1", channel_id="ch1", message_id="m1", limit=51
        )


@pytest.mark.asyncio
async def test_list_channel_messages_page_token_uses_with_url():
    # Covers the shared _paged page-token branch used by all four channel tools:
    # a page_token must decode to a URL and route through builder.with_url(url).get(),
    # not the first-page builder.get(request_configuration=...) path.
    from outlook_mcp.graph.pagination import encode_next_link

    graph = MagicMock()
    with_url = MagicMock()
    with_url.get = AsyncMock(return_value=_fake_collection([_fake_msg()]))
    messages = MagicMock()
    messages.with_url = MagicMock(return_value=with_url)
    channel = MagicMock(messages=messages)
    team = MagicMock()
    team.channels.by_channel_id = MagicMock(return_value=channel)
    graph.raw.teams.by_team_id = MagicMock(return_value=team)

    url = "https://graph.microsoft.com/v1.0/teams/t1/channels/ch1/messages?$skip=25"
    token = encode_next_link(url)

    result = await teams_channels.list_channel_messages(
        graph=graph, team_id="t1", channel_id="ch1", page_token=token
    )

    messages.with_url.assert_called_once_with(url)
    messages.get.assert_not_called()
    assert result["items"][0]["id"] == "m1"
