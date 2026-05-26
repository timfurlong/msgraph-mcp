"""Outlook inbox mail-rule tools.

Wraps the Graph endpoint /me/mailFolders/inbox/messageRules. All rule
operations target the inbox folder — Outlook doesn't support per-folder
rules (the official endpoint is hardcoded to the inbox).
"""

from __future__ import annotations

from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.message_rule import MessageRule
from msgraph.generated.models.message_rule_actions import MessageRuleActions
from msgraph.generated.models.message_rule_predicates import MessageRulePredicates
from msgraph.generated.models.recipient import Recipient

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphValidationError, map_kiota_error
from outlook_mcp.graph.serialize import message_rule_to_dict
from outlook_mcp.graph.trimming import trim_message_rule


def _rules_endpoint(graph, mailbox: str | None):
    return (
        graph.mailbox(mailbox)
        .mail_folders.by_mail_folder_id("inbox")
        .message_rules
    )


async def list_rules(*, graph, mailbox: str | None = None, include_raw: bool = False) -> dict:
    """List inbox mail rules.

    Returns:
        {"items": [trimmed_rule, ...], "next_page_token": None}
    """
    try:
        collection = await _rules_endpoint(graph, mailbox).get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    items = [
        trim_message_rule(message_rule_to_dict(r), include_raw=include_raw)
        for r in (collection.value or [])
    ]
    return {"items": items, "next_page_token": None}


async def get_rule(
    *, graph, rule_id: str, mailbox: str | None = None, include_raw: bool = False
) -> dict:
    """Get a single inbox mail rule by id.

    Args:
        rule_id: Graph rule id. Required.
        mailbox: Optional mailbox (email or user id) for shared mailboxes.
        include_raw: Include the raw Graph payload.

    Returns:
        Trimmed rule dict.
    """
    if not rule_id:
        raise GraphValidationError("`rule_id` is required")
    try:
        rule = await _rules_endpoint(graph, mailbox).by_message_rule_id(rule_id).get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message_rule(message_rule_to_dict(rule), include_raw=include_raw)


def _recipients_from_addresses(addresses: list[str] | None) -> list[Recipient] | None:
    if not addresses:
        return None
    out: list[Recipient] = []
    for addr in addresses:
        if not addr or "@" not in addr:
            raise GraphValidationError(f"Invalid email address: {addr!r}")
        email = EmailAddress()
        email.address = addr
        recipient = Recipient()
        recipient.email_address = email
        out.append(recipient)
    return out


def _build_predicates(
    *,
    sender_contains: list[str] | None,
    subject_contains: list[str] | None,
    body_contains: list[str] | None,
    body_or_subject_contains: list[str] | None,
    from_addresses: list[str] | None,
    has_attachments: bool | None,
) -> MessageRulePredicates | None:
    # Only assign attrs that have real values. Setting `pred.x = None` registers
    # as an explicit null in kiota's backing store and the SDK then emits it as
    # `write_null_value(...)` on the *parent* writer during serialization, which
    # mixes scalar nulls with object keys and trips "Invalid Json output".
    from_recipients = _recipients_from_addresses(from_addresses)
    fields: dict = {}
    if sender_contains:
        fields["sender_contains"] = sender_contains
    if subject_contains:
        fields["subject_contains"] = subject_contains
    if body_contains:
        fields["body_contains"] = body_contains
    if body_or_subject_contains:
        fields["body_or_subject_contains"] = body_or_subject_contains
    if from_recipients:
        fields["from_addresses"] = from_recipients
    if has_attachments is not None:
        fields["has_attachments"] = has_attachments
    if not fields:
        return None
    pred = MessageRulePredicates()
    for name, value in fields.items():
        setattr(pred, name, value)
    return pred


def _build_actions(
    *,
    move_to_folder: str | None,
    mark_as_read: bool | None,
    delete: bool | None,
    stop_processing_rules: bool | None,
) -> MessageRuleActions | None:
    # See note in _build_predicates: only set non-None attrs to avoid the
    # backing-store "explicit null" serialization bug.
    fields: dict = {}
    if move_to_folder is not None:
        fields["move_to_folder"] = move_to_folder
    if mark_as_read is not None:
        fields["mark_as_read"] = mark_as_read
    if delete is not None:
        fields["delete"] = delete
    if stop_processing_rules is not None:
        fields["stop_processing_rules"] = stop_processing_rules
    if not fields:
        return None
    actions = MessageRuleActions()
    for name, value in fields.items():
        setattr(actions, name, value)
    return actions


async def create_rule(
    *,
    graph,
    display_name: str,
    sender_contains: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    body_or_subject_contains: list[str] | None = None,
    from_addresses: list[str] | None = None,
    has_attachments: bool | None = None,
    move_to_folder: str | None = None,
    mark_as_read: bool | None = None,
    delete: bool | None = None,
    stop_processing_rules: bool | None = None,
    sequence: int | None = None,
    is_enabled: bool = True,
    mailbox: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Create an inbox mail rule.

    At least one condition AND one action are required. Conditions
    within a rule are AND-ed by Outlook. Pass multiple values in a list
    (e.g. ``sender_contains=["example.com", "monitor.io"]``) for OR
    within that condition.

    Common condition args:
        sender_contains: Substrings to match against the sender display
            name or address (e.g. ["example.com"]).
        subject_contains: Substrings to match against the subject.
        body_contains: Substrings to match against the message body.
        body_or_subject_contains: Match either subject or body.
        from_addresses: Exact email addresses to match (e.g.
            ["noreply@example.com"]).
        has_attachments: True/False to require/exclude attachments.

    Common action args:
        move_to_folder: Destination folder id (use list_folders or
            create_folder to get one).
        mark_as_read: Mark matching messages as read.
        delete: Move matching messages to Deleted Items.
        stop_processing_rules: If True, no further rules run after this
            one matches. Recommended for routing-to-folder rules.

    Other args:
        sequence: Optional rule order (lower runs first).
        is_enabled: Whether the rule is active. Default True.
        mailbox: Optional mailbox for shared mailboxes.
        include_raw: Include the raw Graph payload.

    Returns:
        The trimmed new rule.
    """
    if not display_name or not display_name.strip():
        raise GraphValidationError("`display_name` is required")
    conditions = _build_predicates(
        sender_contains=sender_contains,
        subject_contains=subject_contains,
        body_contains=body_contains,
        body_or_subject_contains=body_or_subject_contains,
        from_addresses=from_addresses,
        has_attachments=has_attachments,
    )
    if conditions is None:
        raise GraphValidationError(
            "At least one condition is required "
            "(sender_contains, subject_contains, body_contains, "
            "body_or_subject_contains, from_addresses, or has_attachments)"
        )
    actions = _build_actions(
        move_to_folder=move_to_folder,
        mark_as_read=mark_as_read,
        delete=delete,
        stop_processing_rules=stop_processing_rules,
    )
    if actions is None:
        raise GraphValidationError(
            "At least one action is required "
            "(move_to_folder, mark_as_read, delete, or stop_processing_rules)"
        )

    body = MessageRule()
    body.display_name = display_name
    body.is_enabled = is_enabled
    body.conditions = conditions
    body.actions = actions
    # Graph rejects sequence=0 ("InvalidValue"). Default to 1 so the rule lands
    # at the head of the list — callers can override. Never set None here, or
    # the backing store will emit `write_null_value("sequence")` on the parent
    # writer (which conflicts with the serialized object dict; see kiota's
    # BackingStoreSerializationWriterProxyFactory).
    body.sequence = sequence if sequence is not None else 1

    try:
        created = await _rules_endpoint(graph, mailbox).post(body)
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_message_rule(message_rule_to_dict(created), include_raw=include_raw)


async def delete_rule(*, graph, rule_id: str, mailbox: str | None = None) -> dict:
    """Delete an inbox mail rule by id.

    Args:
        rule_id: Graph rule id. Required.
        mailbox: Optional mailbox (email or user id) for shared mailboxes.

    Returns:
        {"deleted": True, "id": <rule_id>}
    """
    if not rule_id:
        raise GraphValidationError("`rule_id` is required")
    try:
        await _rules_endpoint(graph, mailbox).by_message_rule_id(rule_id).delete()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return {"deleted": True, "id": rule_id}


def register(mcp, *, graph) -> None:
    @mcp.tool(name="list_rules", description=list_rules.__doc__ or "")
    async def _list_rules(mailbox: str | None = None, include_raw: bool = False):
        return await list_rules(graph=graph, mailbox=mailbox, include_raw=include_raw)

    @mcp.tool(name="get_rule", description=get_rule.__doc__ or "")
    async def _get_rule(
        rule_id: str, mailbox: str | None = None, include_raw: bool = False
    ):
        return await get_rule(
            graph=graph, rule_id=rule_id, mailbox=mailbox, include_raw=include_raw
        )

    @mcp.tool(name="create_rule", description=create_rule.__doc__ or "")
    async def _create_rule(
        display_name: str,
        sender_contains: list[str] | None = None,
        subject_contains: list[str] | None = None,
        body_contains: list[str] | None = None,
        body_or_subject_contains: list[str] | None = None,
        from_addresses: list[str] | None = None,
        has_attachments: bool | None = None,
        move_to_folder: str | None = None,
        mark_as_read: bool | None = None,
        delete: bool | None = None,
        stop_processing_rules: bool | None = None,
        sequence: int | None = None,
        is_enabled: bool = True,
        mailbox: str | None = None,
        include_raw: bool = False,
    ):
        return await create_rule(
            graph=graph,
            display_name=display_name,
            sender_contains=sender_contains,
            subject_contains=subject_contains,
            body_contains=body_contains,
            body_or_subject_contains=body_or_subject_contains,
            from_addresses=from_addresses,
            has_attachments=has_attachments,
            move_to_folder=move_to_folder,
            mark_as_read=mark_as_read,
            delete=delete,
            stop_processing_rules=stop_processing_rules,
            sequence=sequence,
            is_enabled=is_enabled,
            mailbox=mailbox,
            include_raw=include_raw,
        )

    @mcp.tool(name="delete_rule", description=delete_rule.__doc__ or "")
    async def _delete_rule(rule_id: str, mailbox: str | None = None):
        return await delete_rule(graph=graph, rule_id=rule_id, mailbox=mailbox)
