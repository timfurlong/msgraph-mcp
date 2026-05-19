"""Mail workflow shortcuts: archive, mark read/unread, flag/unflag.

These are thin wrappers around move_message / update_message but with
distinct tool names so the agent picks the right one for natural-language
intents like 'archive this' or 'mark as read'.
"""

from __future__ import annotations

from outlook_mcp.tools.mail_folders import move_message
from outlook_mcp.tools.mail_write import update_message


async def archive_message(
    *, graph, message_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Archive a message (move it to the Archive folder)."""
    return await move_message(
        graph=graph, message_id=message_id, destination="archive",
        mailbox=mailbox, include_raw=include_raw,
    )


async def mark_read(
    *, graph, message_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Mark a message as read."""
    return await update_message(
        graph=graph, message_id=message_id, is_read=True,
        mailbox=mailbox, include_raw=include_raw,
    )


async def mark_unread(
    *, graph, message_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Mark a message as unread."""
    return await update_message(
        graph=graph, message_id=message_id, is_read=False,
        mailbox=mailbox, include_raw=include_raw,
    )


async def flag_message(
    *, graph, message_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Set the follow-up flag on a message."""
    return await update_message(
        graph=graph, message_id=message_id, flag="flagged",
        mailbox=mailbox, include_raw=include_raw,
    )


async def unflag_message(
    *, graph, message_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Clear the follow-up flag on a message."""
    return await update_message(
        graph=graph, message_id=message_id, flag="notFlagged",
        mailbox=mailbox, include_raw=include_raw,
    )


def register(mcp, *, graph) -> None:
    @mcp.tool(name="archive_message", description=archive_message.__doc__ or "")
    async def _archive(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await archive_message(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="mark_read", description=mark_read.__doc__ or "")
    async def _mark_read(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await mark_read(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="mark_unread", description=mark_unread.__doc__ or "")
    async def _mark_unread(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await mark_unread(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="flag_message", description=flag_message.__doc__ or "")
    async def _flag(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await flag_message(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="unflag_message", description=unflag_message.__doc__ or "")
    async def _unflag(message_id: str, mailbox: str | None = None, include_raw: bool = False):
        return await unflag_message(graph=graph, message_id=message_id, mailbox=mailbox, include_raw=include_raw)
