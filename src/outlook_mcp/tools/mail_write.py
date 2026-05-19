"""Mail write tools: send, draft, reply, reply_all, forward, update, delete."""

from __future__ import annotations

import base64
from typing import Literal

from msgraph.generated.models.attachment import Attachment
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.models.followup_flag import FollowupFlag
from msgraph.generated.models.followup_flag_status import FollowupFlagStatus
from msgraph.generated.models.importance import Importance
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.users.item.messages.item.forward.forward_post_request_body import (
    ForwardPostRequestBody,
)
from msgraph.generated.users.item.messages.item.reply.reply_post_request_body import (
    ReplyPostRequestBody,
)
from msgraph.generated.users.item.messages.item.reply_all.reply_all_post_request_body import (
    ReplyAllPostRequestBody,
)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphValidationError, map_kiota_error
from outlook_mcp.graph.serialize import message_to_dict
from outlook_mcp.graph.trimming import trim_message


_MAX_INLINE_ATTACHMENT_BYTES = 3 * 1024 * 1024


def _recipient(addr: str) -> Recipient:
    r = Recipient()
    e = EmailAddress()
    e.address = addr
    r.email_address = e
    return r


def _recipients(addrs: list[str] | None) -> list[Recipient]:
    return [_recipient(a) for a in (addrs or [])]


def _body(text: str, body_type: Literal["text", "html"]) -> ItemBody:
    b = ItemBody()
    b.content_type = BodyType.Text if body_type == "text" else BodyType.Html
    b.content = text
    return b


def _validate_and_build_attachments(atts: list[dict] | None) -> list[Attachment]:
    if not atts:
        return []
    total = 0
    built: list[Attachment] = []
    for spec in atts:
        try:
            raw = base64.b64decode(spec["content_b64"], validate=True)
        except (KeyError, ValueError) as exc:
            raise GraphValidationError(f"Invalid attachment content_b64: {exc}") from exc
        total += len(raw)
        if total > _MAX_INLINE_ATTACHMENT_BYTES:
            raise GraphValidationError(
                f"Attachment payload exceeds 3 MB raw limit (got {total} bytes). "
                "Use a chunked upload (not supported in v1) for larger files."
            )
        fa = FileAttachment()
        fa.name = spec.get("name") or "attachment"
        fa.content_type = spec.get("content_type") or "application/octet-stream"
        fa.content_bytes = raw
        fa.odata_type = "#microsoft.graph.fileAttachment"
        built.append(fa)
    return built


def _build_message(
    *,
    to: list[str],
    subject: str,
    body: str,
    body_type: Literal["text", "html"],
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[dict] | None = None,
) -> Message:
    m = Message()
    m.subject = subject
    m.body = _body(body, body_type)
    m.to_recipients = _recipients(to)
    m.cc_recipients = _recipients(cc)
    m.bcc_recipients = _recipients(bcc)
    atts = _validate_and_build_attachments(attachments)
    if atts:
        m.attachments = atts
    return m


async def send_message(
    *,
    graph,
    to: list[str],
    subject: str,
    body: str,
    body_type: Literal["text", "html"] = "text",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[dict] | None = None,
    save_to_sent_items: bool = True,
    mailbox: str | None = None,
) -> dict:
    """Send a new message immediately.

    Args:
        to: List of recipient email addresses (required, non-empty).
        subject: Subject line.
        body: Message body.
        body_type: "text" or "html". Default "text".
        cc: Optional CC list.
        bcc: Optional BCC list.
        attachments: Optional list of {name, content_b64, content_type}.
            Total raw size must be <= 3 MB.
        save_to_sent_items: Default True.
        mailbox: Optional mailbox (defaults to signed-in user).

    Returns:
        {"status": "sent"}
    """
    if not to:
        raise GraphValidationError("`to` must be a non-empty list of email addresses")
    msg = _build_message(
        to=to, subject=subject, body=body, body_type=body_type,
        cc=cc, bcc=bcc, attachments=attachments,
    )
    req = SendMailPostRequestBody()
    req.message = msg
    req.save_to_sent_items = save_to_sent_items
    try:
        await graph.mailbox(mailbox).send_mail.post(req)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "sent"}


async def create_draft(
    *,
    graph,
    to: list[str],
    subject: str,
    body: str,
    body_type: Literal["text", "html"] = "text",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[dict] | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Create a draft message (not sent). Returns the trimmed draft."""
    msg = _build_message(
        to=to, subject=subject, body=body, body_type=body_type,
        cc=cc, bcc=bcc, attachments=attachments,
    )
    try:
        created = await graph.mailbox(mailbox).messages.post(msg)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message(message_to_dict(created), include_body=False, include_raw=include_raw)


def _reply_body(comment: str | None, to: list[str] | None) -> ReplyPostRequestBody:
    body = ReplyPostRequestBody()
    if comment is not None:
        body.comment = comment
    if to:
        m = Message()
        m.to_recipients = _recipients(to)
        body.message = m
    return body


def _reply_all_body(comment: str | None) -> ReplyAllPostRequestBody:
    body = ReplyAllPostRequestBody()
    if comment is not None:
        body.comment = comment
    return body


def _forward_body(comment: str | None, to: list[str]) -> ForwardPostRequestBody:
    body = ForwardPostRequestBody()
    if comment is not None:
        body.comment = comment
    body.to_recipients = _recipients(to)
    return body


async def reply_message(
    *,
    graph,
    message_id: str,
    comment: str | None = None,
    extra_to: list[str] | None = None,
    mailbox: str | None = None,
) -> dict:
    """Reply to the sender of a message."""
    try:
        await (
            graph.mailbox(mailbox)
            .messages.by_message_id(message_id)
            .reply.post(_reply_body(comment, extra_to))
        )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "sent"}


async def reply_all_message(
    *,
    graph,
    message_id: str,
    comment: str | None = None,
    mailbox: str | None = None,
) -> dict:
    """Reply-all to a message (sender + everyone on the to/cc lines)."""
    try:
        await (
            graph.mailbox(mailbox)
            .messages.by_message_id(message_id)
            .reply_all.post(_reply_all_body(comment))
        )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "sent"}


async def forward_message(
    *,
    graph,
    message_id: str,
    to: list[str],
    comment: str | None = None,
    mailbox: str | None = None,
) -> dict:
    """Forward a message to new recipients."""
    if not to:
        raise GraphValidationError("`to` must be a non-empty list of email addresses")
    try:
        await (
            graph.mailbox(mailbox)
            .messages.by_message_id(message_id)
            .forward.post(_forward_body(comment, to))
        )
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "sent"}


_IMPORTANCE_MAP = {"low": Importance.Low, "normal": Importance.Normal, "high": Importance.High}
_FLAG_MAP = {
    "notFlagged": FollowupFlagStatus.NotFlagged,
    "flagged": FollowupFlagStatus.Flagged,
    "complete": FollowupFlagStatus.Complete,
}


async def update_message(
    *,
    graph,
    message_id: str,
    is_read: bool | None = None,
    flag: Literal["notFlagged", "flagged", "complete"] | None = None,
    importance: Literal["low", "normal", "high"] | None = None,
    categories: list[str] | None = None,
    parent_folder_id: str | None = None,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Patch a message's read state / flag / importance / categories.

    For moving a message between folders, prefer `move_message` (clearer
    intent). This tool patches in-place; setting parent_folder_id here
    works but is not the primary path.

    Returns the trimmed updated message.
    """
    patch = Message()
    if is_read is not None:
        patch.is_read = is_read
    if flag is not None:
        if flag not in _FLAG_MAP:
            raise GraphValidationError(f"flag must be one of {list(_FLAG_MAP)}")
        ff = FollowupFlag()
        ff.flag_status = _FLAG_MAP[flag]
        patch.flag = ff
    if importance is not None:
        if importance not in _IMPORTANCE_MAP:
            raise GraphValidationError(f"importance must be one of {list(_IMPORTANCE_MAP)}")
        patch.importance = _IMPORTANCE_MAP[importance]
    if categories is not None:
        patch.categories = list(categories)
    if parent_folder_id is not None:
        patch.parent_folder_id = parent_folder_id

    try:
        updated = await graph.mailbox(mailbox).messages.by_message_id(message_id).patch(patch)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message(message_to_dict(updated), include_body=False, include_raw=include_raw)


async def delete_message(
    *, graph, message_id: str, mailbox: str | None = None
) -> dict:
    """Soft-delete a message (moves it to Deleted Items)."""
    try:
        await graph.mailbox(mailbox).messages.by_message_id(message_id).delete()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"status": "deleted"}


def register(mcp, *, graph) -> None:
    @mcp.tool(name="send_message", description=send_message.__doc__ or "")
    async def _send(
        to: list[str], subject: str, body: str,
        body_type: Literal["text", "html"] = "text",
        cc: list[str] | None = None,
        bcc: list[str] | None = None, attachments: list[dict] | None = None,
        save_to_sent_items: bool = True, mailbox: str | None = None,
    ):
        return await send_message(
            graph=graph, to=to, subject=subject, body=body,
            body_type=body_type, cc=cc, bcc=bcc, attachments=attachments,
            save_to_sent_items=save_to_sent_items, mailbox=mailbox,
        )

    @mcp.tool(name="create_draft", description=create_draft.__doc__ or "")
    async def _draft(
        to: list[str], subject: str, body: str,
        body_type: Literal["text", "html"] = "text",
        cc: list[str] | None = None,
        bcc: list[str] | None = None, attachments: list[dict] | None = None,
        mailbox: str | None = None, include_raw: bool = False,
    ):
        return await create_draft(
            graph=graph, to=to, subject=subject, body=body,
            body_type=body_type, cc=cc, bcc=bcc, attachments=attachments,
            mailbox=mailbox, include_raw=include_raw,
        )

    @mcp.tool(name="reply_message", description=reply_message.__doc__ or "")
    async def _reply(message_id: str, comment: str | None = None, extra_to: list[str] | None = None, mailbox: str | None = None):
        return await reply_message(graph=graph, message_id=message_id, comment=comment, extra_to=extra_to, mailbox=mailbox)

    @mcp.tool(name="reply_all_message", description=reply_all_message.__doc__ or "")
    async def _reply_all(message_id: str, comment: str | None = None, mailbox: str | None = None):
        return await reply_all_message(graph=graph, message_id=message_id, comment=comment, mailbox=mailbox)

    @mcp.tool(name="forward_message", description=forward_message.__doc__ or "")
    async def _forward(message_id: str, to: list[str], comment: str | None = None, mailbox: str | None = None):
        return await forward_message(graph=graph, message_id=message_id, to=to, comment=comment, mailbox=mailbox)

    @mcp.tool(name="update_message", description=update_message.__doc__ or "")
    async def _update(
        message_id: str,
        is_read: bool | None = None,
        flag: Literal["notFlagged", "flagged", "complete"] | None = None,
        importance: Literal["low", "normal", "high"] | None = None,
        categories: list[str] | None = None,
        parent_folder_id: str | None = None,
        mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await update_message(
            graph=graph, message_id=message_id,
            is_read=is_read, flag=flag, importance=importance,
            categories=categories, parent_folder_id=parent_folder_id,
            mailbox=mailbox, include_raw=include_raw,
        )

    @mcp.tool(name="delete_message", description=delete_message.__doc__ or "")
    async def _delete(message_id: str, mailbox: str | None = None):
        return await delete_message(graph=graph, message_id=message_id, mailbox=mailbox)
