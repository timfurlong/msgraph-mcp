"""Teams channel read tools: joined teams, channels, channel messages, replies."""

from __future__ import annotations

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.teams.item.channels.channels_request_builder import (
    ChannelsRequestBuilder,
)
from msgraph.generated.teams.item.channels.item.messages.messages_request_builder import (
    MessagesRequestBuilder as ChannelMessagesRequestBuilder,
)
from msgraph.generated.teams.item.channels.item.messages.item.replies.replies_request_builder import (
    RepliesRequestBuilder,
)
from msgraph.generated.users.item.joined_teams.joined_teams_request_builder import (
    JoinedTeamsRequestBuilder,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import map_kiota_error
from outlook_mcp.graph.pagination import decode_page_token, encode_next_link, validate_limit
from outlook_mcp.graph.serialize import channel_to_dict, chat_message_to_dict, team_to_dict
from outlook_mcp.graph.trimming import trim_channel, trim_chat_message, trim_team

_CHANNEL_MSG_MAX = 50


def _joined_teams_query(*, limit: int):
    qp = JoinedTeamsRequestBuilder.JoinedTeamsRequestBuilderGetQueryParameters(top=limit)
    return RequestConfiguration[
        JoinedTeamsRequestBuilder.JoinedTeamsRequestBuilderGetQueryParameters
    ](query_parameters=qp)


def _channels_query(*, limit: int):
    qp = ChannelsRequestBuilder.ChannelsRequestBuilderGetQueryParameters(top=limit)
    return RequestConfiguration[
        ChannelsRequestBuilder.ChannelsRequestBuilderGetQueryParameters
    ](query_parameters=qp)


def _channel_messages_query(*, limit: int):
    qp = ChannelMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(top=limit)
    return RequestConfiguration[
        ChannelMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters
    ](query_parameters=qp)


def _replies_query(*, limit: int):
    qp = RepliesRequestBuilder.RepliesRequestBuilderGetQueryParameters(top=limit)
    return RequestConfiguration[
        RepliesRequestBuilder.RepliesRequestBuilderGetQueryParameters
    ](query_parameters=qp)


async def _paged(builder, *, page_token, request_configuration):
    if page_token is not None:
        url = decode_page_token(page_token)
        return await builder.with_url(url).get()
    return await builder.get(request_configuration=request_configuration)


async def list_joined_teams(
    *, graph, limit: int = 25, page_token: str | None = None, include_raw: bool = False
) -> dict:
    """List the teams the signed-in user is a member of.

    Args:
        limit: 1-100. Default 25.
        page_token: Continuation token from a previous result.
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_team, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit)
    builder = graph.raw.me.joined_teams
    try:
        collection = await _paged(builder, page_token=page_token, request_configuration=_joined_teams_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [trim_team(team_to_dict(t), include_raw=include_raw) for t in (collection.value or [])]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def list_channels(
    *, graph, team_id: str, limit: int = 25, page_token: str | None = None, include_raw: bool = False
) -> dict:
    """List channels in a team.

    Args:
        team_id: Graph team id (from list_joined_teams).
        limit: 1-100. Default 25.
        page_token: Continuation token from a previous result.
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_channel, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit)
    builder = graph.raw.teams.by_team_id(team_id).channels
    try:
        collection = await _paged(builder, page_token=page_token, request_configuration=_channels_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [trim_channel(channel_to_dict(c), include_raw=include_raw) for c in (collection.value or [])]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def list_channel_messages(
    *,
    graph,
    team_id: str,
    channel_id: str,
    limit: int = 25,
    page_token: str | None = None,
    include_body: bool = False,
    include_raw: bool = False,
) -> dict:
    """List root messages in a channel, newest first.

    Replies are not included here; fetch them with list_message_replies.

    Args:
        team_id: Graph team id.
        channel_id: Graph channel id (from list_channels).
        limit: 1-50 (Graph caps channel message pages at 50). Default 25.
        page_token: Continuation token from a previous result.
        include_body: When True, include each message's full body. Default
            False (snippet only).
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_chat_message, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit, maximum=_CHANNEL_MSG_MAX)
    builder = graph.raw.teams.by_team_id(team_id).channels.by_channel_id(channel_id).messages
    try:
        collection = await _paged(builder, page_token=page_token, request_configuration=_channel_messages_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [
        trim_chat_message(chat_message_to_dict(m), include_body=include_body, include_raw=include_raw)
        for m in (collection.value or [])
    ]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def list_message_replies(
    *,
    graph,
    team_id: str,
    channel_id: str,
    message_id: str,
    limit: int = 25,
    page_token: str | None = None,
    include_body: bool = False,
    include_raw: bool = False,
) -> dict:
    """List replies to a channel message (the thread under a root post), newest first.

    Args:
        team_id: Graph team id.
        channel_id: Graph channel id.
        message_id: Graph id of the root channel message (from list_channel_messages).
        limit: 1-50 (Graph caps reply pages at 50). Default 25.
        page_token: Continuation token from a previous result.
        include_body: When True, include each reply's full body. Default False.
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_chat_message, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit, maximum=_CHANNEL_MSG_MAX)
    builder = (
        graph.raw.teams.by_team_id(team_id)
        .channels.by_channel_id(channel_id)
        .messages.by_chat_message_id(message_id)
        .replies
    )
    try:
        collection = await _paged(builder, page_token=page_token, request_configuration=_replies_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [
        trim_chat_message(chat_message_to_dict(m), include_body=include_body, include_raw=include_raw)
        for m in (collection.value or [])
    ]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


def register(mcp, *, graph) -> None:
    @mcp.tool(name="list_joined_teams", description=list_joined_teams.__doc__ or "")
    async def _list_joined_teams(limit: int = 25, page_token: str | None = None, include_raw: bool = False):
        return await list_joined_teams(graph=graph, limit=limit, page_token=page_token, include_raw=include_raw)

    @mcp.tool(name="list_channels", description=list_channels.__doc__ or "")
    async def _list_channels(team_id: str, limit: int = 25, page_token: str | None = None, include_raw: bool = False):
        return await list_channels(graph=graph, team_id=team_id, limit=limit, page_token=page_token, include_raw=include_raw)

    @mcp.tool(name="list_channel_messages", description=list_channel_messages.__doc__ or "")
    async def _list_channel_messages(
        team_id: str,
        channel_id: str,
        limit: int = 25,
        page_token: str | None = None,
        include_body: bool = False,
        include_raw: bool = False,
    ):
        return await list_channel_messages(
            graph=graph, team_id=team_id, channel_id=channel_id, limit=limit,
            page_token=page_token, include_body=include_body, include_raw=include_raw,
        )

    @mcp.tool(name="list_message_replies", description=list_message_replies.__doc__ or "")
    async def _list_message_replies(
        team_id: str,
        channel_id: str,
        message_id: str,
        limit: int = 25,
        page_token: str | None = None,
        include_body: bool = False,
        include_raw: bool = False,
    ):
        return await list_message_replies(
            graph=graph, team_id=team_id, channel_id=channel_id, message_id=message_id,
            limit=limit, page_token=page_token, include_body=include_body, include_raw=include_raw,
        )
