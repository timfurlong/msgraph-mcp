"""Mail folder + move tools."""

from __future__ import annotations

from msgraph.generated.users.item.messages.item.move.move_post_request_body import (
    MovePostRequestBody,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphValidationError, map_kiota_error
from outlook_mcp.graph.serialize import folder_to_dict, message_to_dict
from outlook_mcp.graph.trimming import trim_folder, trim_message


async def list_folders(
    *,
    graph,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """List the user's mail folders (top level).

    Returns:
        {"items": [trimmed_folder, ...], "next_page_token": None}
    """
    try:
        collection = await graph.mailbox(mailbox).mail_folders.get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [trim_folder(folder_to_dict(f), include_raw=include_raw) for f in (collection.value or [])]
    return {"items": items, "next_page_token": None}


async def move_message(
    *,
    graph,
    message_id: str,
    destination: str,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Move a message to another folder.

    Args:
        message_id: Graph message id.
        destination: Target folder id, OR one of the well-known names:
            'archive', 'inbox', 'junkemail', 'deleteditems', 'sentitems', 'drafts'.
        mailbox: Optional mailbox.
        include_raw: Include the raw Graph payload.

    Returns:
        The trimmed moved message (which lives in the destination folder).
    """
    if not destination:
        raise GraphValidationError("`destination` is required")
    body = MovePostRequestBody()
    body.destination_id = destination
    try:
        moved = await (
            graph.mailbox(mailbox).messages.by_message_id(message_id).move.post(body)
        )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message(message_to_dict(moved), include_body=False, include_raw=include_raw)


def register(mcp, *, graph) -> None:
    @mcp.tool(name="list_folders", description=list_folders.__doc__ or "")
    async def _list_folders(mailbox: str | None = None, include_raw: bool = False):
        return await list_folders(graph=graph, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="move_message", description=move_message.__doc__ or "")
    async def _move_message(
        message_id: str, destination: str, mailbox: str | None = None, include_raw: bool = False
    ):
        return await move_message(
            graph=graph, message_id=message_id, destination=destination,
            mailbox=mailbox, include_raw=include_raw,
        )
