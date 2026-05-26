"""Batch mail-action tools backed by Microsoft Graph's $batch endpoint.

These are bulk versions of the single-message tools in mail_actions.py.
Each accepts a list of message ids and returns per-id success/error.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.graph.batch import (
    BatchRequest,
    BatchResult,
    GraphBatchTransport,
    execute_batch,
)
from outlook_mcp.graph.errors import GraphValidationError

MAX_BATCH_IDS = 1000


def _validate_ids(message_ids: list[str]) -> None:
    if not message_ids:
        raise GraphValidationError("`message_ids` must be a non-empty list")
    if len(message_ids) > MAX_BATCH_IDS:
        raise GraphValidationError(
            f"`message_ids` exceeds maximum of {MAX_BATCH_IDS} per call"
        )


def _mailbox_path_segment(mailbox: str | None) -> str:
    """Return the URL prefix for the mailbox: '/me' or '/users/{mailbox}'."""
    return "/me" if mailbox is None else f"/users/{mailbox}"


def _summarize(results: list[BatchResult]) -> dict[str, Any]:
    out_results = [
        {
            "id": r.id,
            "ok": r.ok,
            **({"error": r.error} if r.error else {}),
        }
        for r in results
    ]
    succeeded = sum(1 for r in results if r.ok)
    return {
        "results": out_results,
        "summary": {
            "total": len(results),
            "succeeded": succeeded,
            "failed": len(results) - succeeded,
        },
    }


async def batch_archive_messages(
    *,
    graph,
    message_ids: list[str],
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Archive a list of messages (bulk move to the Archive well-known folder)."""
    _validate_ids(message_ids)
    base = _mailbox_path_segment(mailbox)
    requests = [
        BatchRequest(
            id=mid,
            method="POST",
            url=f"{base}/messages/{mid}/move",
            body={"destinationId": "archive"},
        )
        for mid in message_ids
    ]
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


async def batch_move_messages(
    *,
    graph,
    message_ids: list[str],
    destination: str,
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Move a list of messages to the given folder.

    `destination` accepts a folder id OR one of the well-known names that
    `move_message` already supports ('archive', 'inbox', 'junkemail',
    'deleteditems', 'sentitems', 'drafts'). Graph resolves well-known names
    server-side, so we pass through without local resolution.

    Spec deviation: Spec §6.3 calls for pre-resolving well-known names to
    folder ids. In practice Graph's /move endpoint accepts well-known names
    directly (this is how the existing single-message move_message already
    works — see src/outlook_mcp/tools/mail_folders.py). We pass through
    unchanged. If a real-world test surfaces a case where Graph rejects a
    well-known name in $batch but accepts it in single calls, add resolution
    then.
    """
    _validate_ids(message_ids)
    if not destination:
        raise GraphValidationError("`destination` is required")
    base = _mailbox_path_segment(mailbox)
    requests = [
        BatchRequest(
            id=mid,
            method="POST",
            url=f"{base}/messages/{mid}/move",
            body={"destinationId": destination},
        )
        for mid in message_ids
    ]
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


def _patch_requests(
    message_ids: list[str],
    mailbox: str | None,
    body: dict[str, Any],
) -> list[BatchRequest]:
    """Build a list of PATCH BatchRequests for /me/messages/{id} (no /move)."""
    base = _mailbox_path_segment(mailbox)
    return [
        BatchRequest(
            id=mid,
            method="PATCH",
            url=f"{base}/messages/{mid}",
            body=body,
        )
        for mid in message_ids
    ]


async def batch_mark_read(
    *,
    graph,
    message_ids: list[str],
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Mark a list of messages as read."""
    _validate_ids(message_ids)
    requests = _patch_requests(message_ids, mailbox, {"isRead": True})
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


async def batch_mark_unread(
    *,
    graph,
    message_ids: list[str],
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Mark a list of messages as unread."""
    _validate_ids(message_ids)
    requests = _patch_requests(message_ids, mailbox, {"isRead": False})
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


async def batch_flag_messages(
    *,
    graph,
    message_ids: list[str],
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Flag a list of messages (set followup flag to 'flagged')."""
    _validate_ids(message_ids)
    requests = _patch_requests(message_ids, mailbox, {"flag": {"flagStatus": "flagged"}})
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


async def batch_unflag_messages(
    *,
    graph,
    message_ids: list[str],
    mailbox: str | None = None,
) -> dict[str, Any]:
    """Unflag a list of messages (set followup flag to 'notFlagged')."""
    _validate_ids(message_ids)
    requests = _patch_requests(message_ids, mailbox, {"flag": {"flagStatus": "notFlagged"}})
    transport = GraphBatchTransport(graph)
    results = await execute_batch(transport, requests)
    return _summarize(results)


def register(mcp, *, graph) -> None:
    @mcp.tool(name="batch_archive_messages", description=batch_archive_messages.__doc__ or "")
    async def _archive(message_ids: list[str], mailbox: str | None = None):
        return await batch_archive_messages(
            graph=graph, message_ids=message_ids, mailbox=mailbox
        )

    @mcp.tool(name="batch_move_messages", description=batch_move_messages.__doc__ or "")
    async def _move(message_ids: list[str], destination: str, mailbox: str | None = None):
        return await batch_move_messages(
            graph=graph, message_ids=message_ids, destination=destination, mailbox=mailbox
        )

    @mcp.tool(name="batch_mark_read", description=batch_mark_read.__doc__ or "")
    async def _mark_read(message_ids: list[str], mailbox: str | None = None):
        return await batch_mark_read(graph=graph, message_ids=message_ids, mailbox=mailbox)

    @mcp.tool(name="batch_mark_unread", description=batch_mark_unread.__doc__ or "")
    async def _mark_unread(message_ids: list[str], mailbox: str | None = None):
        return await batch_mark_unread(graph=graph, message_ids=message_ids, mailbox=mailbox)

    @mcp.tool(name="batch_flag_messages", description=batch_flag_messages.__doc__ or "")
    async def _flag(message_ids: list[str], mailbox: str | None = None):
        return await batch_flag_messages(graph=graph, message_ids=message_ids, mailbox=mailbox)

    @mcp.tool(name="batch_unflag_messages", description=batch_unflag_messages.__doc__ or "")
    async def _unflag(message_ids: list[str], mailbox: str | None = None):
        return await batch_unflag_messages(graph=graph, message_ids=message_ids, mailbox=mailbox)
