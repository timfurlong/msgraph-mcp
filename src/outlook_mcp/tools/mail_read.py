"""Mail read tools: list, search, get, list/download attachments."""

from __future__ import annotations

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder as FolderMessagesRequestBuilder,
)
from msgraph.generated.users.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import map_kiota_error
from outlook_mcp.graph.pagination import (
    decode_page_token,
    encode_next_link,
    validate_limit,
)
from outlook_mcp.graph.serialize import (
    attachment_to_dict,
    message_to_dict,
)
from outlook_mcp.graph.trimming import (
    trim_attachment_download,
    trim_attachment_list,
    trim_message,
)


def _list_messages_query(*, limit: int):
    qp = FolderMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=limit,
        orderby=["receivedDateTime DESC"],
    )
    return RequestConfiguration[
        FolderMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters
    ](query_parameters=qp)


def _search_query(*, limit: int, query: str):
    qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=limit,
        search=query,
    )
    return RequestConfiguration[
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters
    ](query_parameters=qp)


async def list_messages(
    *,
    graph,
    folder_id: str = "inbox",
    mailbox: str | None = None,
    limit: int = 25,
    page_token: str | None = None,
    include_raw: bool = False,
) -> dict:
    """List messages from a mail folder, newest first.

    Args:
        folder_id: Folder id, or a well-known name (inbox, sentitems, drafts,
            deleteditems, archive, junkemail). Default: inbox.
        mailbox: Optional mailbox (email or user ID). Default: signed-in user.
        limit: 1-100. Default 25.
        page_token: Pass next_page_token from a previous result to continue.
        include_raw: Include the raw Graph payload under "raw" on each item.

    Returns:
        {"items": [trimmed_message, ...], "next_page_token": str | None}
    """
    limit = validate_limit(limit)
    builder = graph.mailbox(mailbox).mail_folders.by_mail_folder_id(folder_id).messages
    try:
        if page_token is not None:
            url = decode_page_token(page_token)
            collection = await builder.with_url(url).get()
        else:
            collection = await builder.get(request_configuration=_list_messages_query(limit=limit))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    items = [
        trim_message(message_to_dict(m), include_body=False, include_raw=include_raw)
        for m in (collection.value or [])
    ]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def search_messages(
    *,
    graph,
    query: str,
    mailbox: str | None = None,
    limit: int = 25,
    page_token: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Search messages across all folders using Graph $search.

    The query is passed to Graph as-is — supply quoting yourself if you need
    a literal phrase (e.g. '"weekly report"') or KQL fielded predicates
    (e.g. 'from:alice subject:"report"').

    Args:
        query: Graph $search expression. Plain tokens match across common
            mail fields; quoted phrases match literally; KQL `field:value`
            forms target specific fields.
        mailbox: Optional mailbox.
        limit: 1-100.
        page_token: Continuation token.
        include_raw: Include raw payloads.
    """
    limit = validate_limit(limit)
    builder = graph.mailbox(mailbox).messages
    try:
        if page_token is not None:
            url = decode_page_token(page_token)
            collection = await builder.with_url(url).get()
        else:
            collection = await builder.get(request_configuration=_search_query(limit=limit, query=query))
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    items = [
        trim_message(message_to_dict(m), include_body=False, include_raw=include_raw)
        for m in (collection.value or [])
    ]
    return {"items": items, "next_page_token": encode_next_link(getattr(collection, "odata_next_link", None))}


async def get_message(
    *,
    graph,
    message_id: str,
    mailbox: str | None = None,
    include_body: bool = False,
    include_raw: bool = False,
) -> dict:
    """Fetch a single message by id.

    Args:
        message_id: Graph message id.
        mailbox: Optional mailbox.
        include_body: When True, the full body is returned. Default False
            (snippet only).
        include_raw: Include the raw Graph payload under "raw".

    Returns:
        Trimmed message object.
    """
    try:
        msg = await graph.mailbox(mailbox).messages.by_message_id(message_id).get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message(message_to_dict(msg), include_body=include_body, include_raw=include_raw)


async def list_attachments(
    *,
    graph,
    message_id: str,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """List attachments on a message (metadata only — no content)."""
    try:
        collection = await graph.mailbox(mailbox).messages.by_message_id(message_id).attachments.get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [
        trim_attachment_list(attachment_to_dict(a, include_content=False), include_raw=include_raw)
        for a in (collection.value or [])
    ]
    return {"items": items, "next_page_token": None}


async def download_attachment(
    *,
    graph,
    message_id: str,
    attachment_id: str,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Download a single attachment as base64.

    Returns: {name, content_type, size_bytes, content_base64}
    """
    try:
        att = await (
            graph.mailbox(mailbox)
            .messages.by_message_id(message_id)
            .attachments.by_attachment_id(attachment_id)
            .get()
        )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_attachment_download(attachment_to_dict(att, include_content=True), include_raw=include_raw)


def register(mcp, *, graph) -> None:
    @mcp.tool(name="list_messages", description=list_messages.__doc__ or "")
    async def _list_messages(
        folder_id: str = "inbox",
        mailbox: str | None = None,
        limit: int = 25,
        page_token: str | None = None,
        include_raw: bool = False,
    ):
        return await list_messages(
            graph=graph, folder_id=folder_id, mailbox=mailbox,
            limit=limit, page_token=page_token, include_raw=include_raw,
        )

    @mcp.tool(name="search_messages", description=search_messages.__doc__ or "")
    async def _search_messages(
        query: str,
        mailbox: str | None = None,
        limit: int = 25,
        page_token: str | None = None,
        include_raw: bool = False,
    ):
        return await search_messages(
            graph=graph, query=query, mailbox=mailbox,
            limit=limit, page_token=page_token, include_raw=include_raw,
        )

    @mcp.tool(name="get_message", description=get_message.__doc__ or "")
    async def _get_message(
        message_id: str,
        mailbox: str | None = None,
        include_body: bool = False,
        include_raw: bool = False,
    ):
        return await get_message(
            graph=graph, message_id=message_id, mailbox=mailbox,
            include_body=include_body, include_raw=include_raw,
        )

    @mcp.tool(name="list_attachments", description=list_attachments.__doc__ or "")
    async def _list_attachments(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await list_attachments(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="download_attachment", description=download_attachment.__doc__ or "")
    async def _download_attachment(
        message_id: str, attachment_id: str, mailbox: str | None = None, include_raw: bool = False
    ):
        return await download_attachment(
            graph=graph, message_id=message_id, attachment_id=attachment_id,
            mailbox=mailbox, include_raw=include_raw,
        )
