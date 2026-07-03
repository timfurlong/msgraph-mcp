from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import teams_content

PNG = b"\x89PNG\r\n\x1a\n\x00\x00\x00\x0dIHDR"


def _wire_chat(graph, data=PNG):
    content = MagicMock()
    content.get = AsyncMock(return_value=data)
    hc_item = MagicMock(content=content)
    hosted = MagicMock()
    hosted.by_chat_message_hosted_content_id = MagicMock(return_value=hc_item)
    msg_item = MagicMock(hosted_contents=hosted)
    messages = MagicMock()
    messages.by_chat_message_id = MagicMock(return_value=msg_item)
    graph.raw.chats.by_chat_id = MagicMock(return_value=MagicMock(messages=messages))
    return content


def _wire_channel(graph, data=PNG):
    content = MagicMock()
    content.get = AsyncMock(return_value=data)
    hc_item = MagicMock(content=content)
    hosted = MagicMock()
    hosted.by_chat_message_hosted_content_id = MagicMock(return_value=hc_item)
    msg_item = MagicMock(hosted_contents=hosted)
    channel = MagicMock()
    channel.messages.by_chat_message_id = MagicMock(return_value=msg_item)
    team = MagicMock()
    team.channels.by_channel_id = MagicMock(return_value=channel)
    graph.raw.teams.by_team_id = MagicMock(return_value=team)
    return content


@pytest.mark.asyncio
async def test_download_from_chat_returns_base64_and_sniffs_png():
    graph = MagicMock()
    _wire_chat(graph)
    result = await teams_content.download_hosted_content(
        graph=graph, chat_id="c1", message_id="m1", hosted_content_id="h1"
    )
    assert result["content_type"] == "image/png"
    assert result["size_bytes"] == len(PNG)
    assert result["content_base64"] is not None


@pytest.mark.asyncio
async def test_download_from_channel_routes_correctly():
    graph = MagicMock()
    _wire_channel(graph)
    result = await teams_content.download_hosted_content(
        graph=graph, team_id="t1", channel_id="ch1", message_id="m1", hosted_content_id="h1"
    )
    graph.raw.teams.by_team_id.assert_called_once_with("t1")
    assert result["content_base64"] is not None


@pytest.mark.asyncio
async def test_download_rejects_no_location():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_content.download_hosted_content(graph=graph, message_id="m1", hosted_content_id="h1")


@pytest.mark.asyncio
async def test_download_rejects_both_locations():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_content.download_hosted_content(
            graph=graph, chat_id="c1", team_id="t1", channel_id="ch1", message_id="m1", hosted_content_id="h1"
        )


@pytest.mark.asyncio
async def test_download_rejects_partial_channel_location():
    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await teams_content.download_hosted_content(
            graph=graph, team_id="t1", message_id="m1", hosted_content_id="h1"
        )
