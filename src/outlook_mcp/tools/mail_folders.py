"""Mail folder + move tools."""

from __future__ import annotations

from msgraph.generated.models.mail_folder import MailFolder
from msgraph.generated.users.item.mail_folders.item.move.move_post_request_body import (
    MovePostRequestBody as FolderMovePostRequestBody,
)
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


async def create_folder(
    *,
    graph,
    display_name: str,
    parent_folder_id: str | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Create a new mail folder (top-level or child).

    Args:
        display_name: Folder name. Required, non-empty.
        parent_folder_id: If set, create as a child of this folder; else
            create at the top level. Accepts a folder id or a well-known
            name ('inbox', 'archive', etc.).
        mailbox: Optional mailbox (email or user id) for shared mailboxes.
        include_raw: Include the raw Graph payload.

    Returns:
        The trimmed new folder.
    """
    if not display_name or not display_name.strip():
        raise GraphValidationError("`display_name` is required")
    body = MailFolder()
    body.display_name = display_name
    try:
        if parent_folder_id:
            created = await (
                graph.mailbox(mailbox)
                .mail_folders.by_mail_folder_id(parent_folder_id)
                .child_folders.post(body)
            )
        else:
            created = await graph.mailbox(mailbox).mail_folders.post(body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_folder(folder_to_dict(created), include_raw=include_raw)


async def update_folder(
    *,
    graph,
    folder_id: str,
    display_name: str | None = None,
    parent_folder_id: str | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Rename and/or reparent a mail folder.

    At least one of ``display_name`` or ``parent_folder_id`` is required.
    Well-known folders (inbox, drafts, sent items, etc.) cannot be
    renamed or moved — Graph will reject the request.

    Args:
        folder_id: Graph folder id. Required.
        display_name: New folder name. None = unchanged.
        parent_folder_id: Move under this folder. Accepts a folder id or
            a well-known name ('inbox', 'archive', etc.). None = unchanged.
        mailbox: Optional mailbox (email or user id) for shared mailboxes.
        include_raw: Include the raw Graph payload.

    Returns:
        The trimmed updated folder.
    """
    if not folder_id:
        raise GraphValidationError("`folder_id` is required")
    if display_name is None and parent_folder_id is None:
        raise GraphValidationError(
            "At least one of `display_name` or `parent_folder_id` is required"
        )
    if display_name is not None and not display_name.strip():
        raise GraphValidationError("`display_name` cannot be empty")

    folder_endpoint = graph.mailbox(mailbox).mail_folders.by_mail_folder_id(folder_id)
    result = None
    try:
        if display_name is not None:
            patch = MailFolder()
            patch.display_name = display_name
            result = await folder_endpoint.patch(patch)
        if parent_folder_id is not None:
            move_body = FolderMovePostRequestBody()
            move_body.destination_id = parent_folder_id
            result = await folder_endpoint.move.post(move_body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_folder(folder_to_dict(result), include_raw=include_raw)


async def delete_folder(
    *,
    graph,
    folder_id: str,
    mailbox: str | None = None,
) -> dict:
    """Delete a mail folder. Refuses to delete folders that aren't empty.

    Reads the folder first and aborts if `totalItemCount > 0` or
    `childFolderCount > 0`. This is a small safety net against accidentally
    nuking a folder with real messages; there is a small race window between
    the read and the delete, so don't rely on this for adversarial safety.

    Args:
        folder_id: Graph folder id. Required.
        mailbox: Optional mailbox (email or user id) for shared mailboxes.

    Returns:
        {"deleted": True, "id": <folder_id>}
    """
    if not folder_id:
        raise GraphValidationError("`folder_id` is required")
    folder_endpoint = graph.mailbox(mailbox).mail_folders.by_mail_folder_id(folder_id)
    try:
        folder = await folder_endpoint.get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    total = getattr(folder, "total_item_count", None) or 0
    children = getattr(folder, "child_folder_count", None) or 0
    if total > 0 or children > 0:
        raise GraphValidationError(
            f"Refusing to delete non-empty folder {folder_id!r}: "
            f"contains {total} message(s) and {children} child folder(s). "
            "Empty it first or move the contents elsewhere."
        )
    try:
        await folder_endpoint.delete()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"deleted": True, "id": folder_id}


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

    @mcp.tool(name="create_folder", description=create_folder.__doc__ or "")
    async def _create_folder(
        display_name: str,
        parent_folder_id: str | None = None,
        mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await create_folder(
            graph=graph,
            display_name=display_name,
            parent_folder_id=parent_folder_id,
            mailbox=mailbox,
            include_raw=include_raw,
        )

    @mcp.tool(name="update_folder", description=update_folder.__doc__ or "")
    async def _update_folder(
        folder_id: str,
        display_name: str | None = None,
        parent_folder_id: str | None = None,
        mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await update_folder(
            graph=graph,
            folder_id=folder_id,
            display_name=display_name,
            parent_folder_id=parent_folder_id,
            mailbox=mailbox,
            include_raw=include_raw,
        )

    @mcp.tool(name="delete_folder", description=delete_folder.__doc__ or "")
    async def _delete_folder(folder_id: str, mailbox: str | None = None):
        return await delete_folder(graph=graph, folder_id=folder_id, mailbox=mailbox)
