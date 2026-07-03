"""Teams chat read tools: list the user's chats and their message history."""

from __future__ import annotations

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.chats.chats_request_builder import ChatsRequestBuilder
from msgraph.generated.chats.item.messages.messages_request_builder import (
    MessagesRequestBuilder as ChatMessagesRequestBuilder,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import map_kiota_error
from outlook_mcp.graph.pagination import decode_page_token, encode_next_link, validate_limit
from outlook_mcp.graph.serialize import chat_message_to_dict, chat_to_dict
from outlook_mcp.graph.trimming import trim_chat, trim_chat_message


def _list_chats_query(*, limit: int):
    qp = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=limit,
        expand=["members"],
    )
    return RequestConfiguration[ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters](
        query_parameters=qp
    )


def _list_chat_messages_query(*, limit: int):
    qp = ChatMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(top=limit)
    return RequestConfiguration[
        ChatMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters
    ](query_parameters=qp)


async def list_chats(
    *,
    graph,
    limit: int = 25,
    page_token: str | None = None,
    include_raw: bool = False,
) -> dict:
    """List the signed-in user's Teams chats (1:1, group, meeting).

    Members are expanded so 1:1 chats (which have no topic) are identifiable
    by participant name. Results come back in Graph's default order.

    Args:
        limit: 1-100. Default 25.
        page_token: Pass next_page_token from a previous result to continue.
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_chat, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit)
    builder = graph.raw.me.chats
    try:
        if page_token is not None:
            url = decode_page_token(page_token)
            collection = await builder.with_url(url).get()
        else:
            collection = await builder.get(request_configuration=_list_chats_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    items = [trim_chat(chat_to_dict(c), include_raw=include_raw) for c in (collection.value or [])]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def list_chat_messages(
    *,
    graph,
    chat_id: str,
    limit: int = 25,
    page_token: str | None = None,
    include_body: bool = False,
    include_raw: bool = False,
) -> dict:
    """List messages in a chat, newest first.

    Args:
        chat_id: Graph chat id (from list_chats).
        limit: 1-100. Default 25.
        page_token: Continuation token from a previous result.
        include_body: When True, include each message's full body. Default
            False (snippet only).
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_chat_message, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit)
    builder = graph.raw.chats.by_chat_id(chat_id).messages
    try:
        if page_token is not None:
            url = decode_page_token(page_token)
            collection = await builder.with_url(url).get()
        else:
            collection = await builder.get(request_configuration=_list_chat_messages_query(limit=limit))
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
    @mcp.tool(name="list_chats", description=list_chats.__doc__ or "")
    async def _list_chats(limit: int = 25, page_token: str | None = None, include_raw: bool = False):
        return await list_chats(graph=graph, limit=limit, page_token=page_token, include_raw=include_raw)

    @mcp.tool(name="list_chat_messages", description=list_chat_messages.__doc__ or "")
    async def _list_chat_messages(
        chat_id: str,
        limit: int = 25,
        page_token: str | None = None,
        include_body: bool = False,
        include_raw: bool = False,
    ):
        return await list_chat_messages(
            graph=graph, chat_id=chat_id, limit=limit,
            page_token=page_token, include_body=include_body, include_raw=include_raw,
        )
