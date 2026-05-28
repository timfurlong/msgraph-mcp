from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools import mail_rules


def _fake_rule(id_="rule-1", name="r", enabled=True):
    return SimpleNamespace(
        id=id_,
        display_name=name,
        sequence=1,
        is_enabled=enabled,
        has_error=False,
        is_read_only=False,
        conditions=None,
        exceptions=None,
        actions=None,
        additional_data={},
    )


def _build_graph_with_rules_endpoint():
    """Wire graph.mailbox(x).mail_folders.by_mail_folder_id('inbox').message_rules
    to a MagicMock so individual tests can attach .get / .post / .by_message_rule_id.
    """
    graph = MagicMock()
    mb = MagicMock()
    inbox = MagicMock()
    mb.mail_folders.by_mail_folder_id = MagicMock(return_value=inbox)
    graph.mailbox = MagicMock(return_value=mb)
    return graph, mb, inbox


@pytest.mark.asyncio
async def test_list_rules_returns_trimmed_items():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    inbox.message_rules.get = AsyncMock(
        return_value=SimpleNamespace(value=[_fake_rule(), _fake_rule(id_="rule-2", name="r2")])
    )

    result = await mail_rules.list_rules(graph=graph)
    assert len(result["items"]) == 2
    assert result["items"][0]["display_name"] == "r"
    assert result["next_page_token"] is None
    mb.mail_folders.by_mail_folder_id.assert_called_once_with("inbox")


@pytest.mark.asyncio
async def test_get_rule_returns_trimmed_rule():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    by_id = MagicMock()
    by_id.get = AsyncMock(return_value=_fake_rule(id_="rule-1", name="Test notifications"))
    inbox.message_rules.by_message_rule_id = MagicMock(return_value=by_id)

    result = await mail_rules.get_rule(graph=graph, rule_id="rule-1")
    assert result["id"] == "rule-1"
    assert result["display_name"] == "Test notifications"
    inbox.message_rules.by_message_rule_id.assert_called_once_with("rule-1")


@pytest.mark.asyncio
async def test_get_rule_requires_rule_id():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.get_rule(graph=graph, rule_id="")


@pytest.mark.asyncio
async def test_create_rule_with_sender_and_move_action():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    created = _fake_rule(id_="new-rule", name="Test notifications")
    inbox.message_rules.post = AsyncMock(return_value=created)

    result = await mail_rules.create_rule(
        graph=graph,
        display_name="Test notifications",
        sender_contains=["example.com"],
        move_to_folder="AAMkFolderId",
        stop_processing_rules=True,
    )
    assert result["id"] == "new-rule"

    posted = inbox.message_rules.post.await_args.args[0]
    assert posted.display_name == "Test notifications"
    assert posted.is_enabled is True
    assert posted.conditions.sender_contains == ["example.com"]
    assert posted.actions.move_to_folder == "AAMkFolderId"
    assert posted.actions.stop_processing_rules is True


@pytest.mark.asyncio
async def test_create_rule_with_from_addresses():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    created = _fake_rule()
    inbox.message_rules.post = AsyncMock(return_value=created)

    await mail_rules.create_rule(
        graph=graph,
        display_name="r",
        from_addresses=["noreply@example.com", "alerts@example.com"],
        move_to_folder="fid",
    )
    posted = inbox.message_rules.post.await_args.args[0]
    addrs = posted.conditions.from_addresses
    assert len(addrs) == 2
    assert addrs[0].email_address.address == "noreply@example.com"
    assert addrs[1].email_address.address == "alerts@example.com"


@pytest.mark.asyncio
async def test_create_rule_requires_display_name():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.create_rule(graph=graph, display_name="", move_to_folder="f")


@pytest.mark.asyncio
async def test_create_rule_requires_at_least_one_condition():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.create_rule(
            graph=graph, display_name="r", move_to_folder="f"
        )


@pytest.mark.asyncio
async def test_create_rule_requires_at_least_one_action():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.create_rule(
            graph=graph, display_name="r", sender_contains=["x"]
        )


@pytest.mark.asyncio
async def test_create_rule_rejects_whitespace_display_name():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.create_rule(
            graph=graph, display_name="   ", sender_contains=["x"], mark_as_read=True
        )


@pytest.mark.asyncio
async def test_create_rule_rejects_invalid_from_address():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError, match="Invalid email address"):
        await mail_rules.create_rule(
            graph=graph,
            display_name="r",
            from_addresses=["not-an-email"],
            move_to_folder="fid",
        )


@pytest.mark.asyncio
async def test_update_rule_top_level_fields_only():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    by_id = MagicMock()
    by_id.patch = AsyncMock(return_value=_fake_rule(id_="rule-1", name="Renamed"))
    inbox.message_rules.by_message_rule_id = MagicMock(return_value=by_id)

    result = await mail_rules.update_rule(
        graph=graph,
        rule_id="rule-1",
        display_name="Renamed",
        is_enabled=False,
        sequence=3,
    )
    assert result["id"] == "rule-1"
    inbox.message_rules.by_message_rule_id.assert_called_once_with("rule-1")
    patched = by_id.patch.await_args.args[0]
    assert patched.display_name == "Renamed"
    assert patched.is_enabled is False
    assert patched.sequence == 3
    # No conditions/actions args passed → blocks left untouched.
    changed = patched.backing_store.enumerate_keys_for_values_changed_to_null()
    assert list(changed) == []


@pytest.mark.asyncio
async def test_update_rule_replaces_conditions_block():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    by_id = MagicMock()
    by_id.patch = AsyncMock(return_value=_fake_rule(id_="rule-1"))
    inbox.message_rules.by_message_rule_id = MagicMock(return_value=by_id)

    await mail_rules.update_rule(
        graph=graph,
        rule_id="rule-1",
        subject_contains=["[urgent]"],
    )
    patched = by_id.patch.await_args.args[0]
    assert patched.conditions is not None
    assert patched.conditions.subject_contains == ["[urgent]"]
    # Actions block not sent → preserved by Graph.
    assert patched.actions is None
    cond_nulls = patched.conditions.backing_store.enumerate_keys_for_values_changed_to_null()
    assert list(cond_nulls) == []


@pytest.mark.asyncio
async def test_update_rule_replaces_actions_block():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    by_id = MagicMock()
    by_id.patch = AsyncMock(return_value=_fake_rule(id_="rule-1"))
    inbox.message_rules.by_message_rule_id = MagicMock(return_value=by_id)

    await mail_rules.update_rule(
        graph=graph,
        rule_id="rule-1",
        move_to_folder="new-folder",
        stop_processing_rules=True,
    )
    patched = by_id.patch.await_args.args[0]
    assert patched.actions is not None
    assert patched.actions.move_to_folder == "new-folder"
    assert patched.actions.stop_processing_rules is True
    assert patched.conditions is None


@pytest.mark.asyncio
async def test_update_rule_requires_rule_id():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.update_rule(graph=graph, rule_id="", display_name="x")


@pytest.mark.asyncio
async def test_update_rule_requires_at_least_one_field():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError, match="At least one field"):
        await mail_rules.update_rule(graph=graph, rule_id="rule-1")


@pytest.mark.asyncio
async def test_update_rule_rejects_whitespace_display_name():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.update_rule(
            graph=graph, rule_id="rule-1", display_name="   "
        )


@pytest.mark.asyncio
async def test_delete_rule_calls_delete_endpoint():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    by_id = MagicMock(delete=AsyncMock(return_value=None))
    inbox.message_rules.by_message_rule_id = MagicMock(return_value=by_id)

    result = await mail_rules.delete_rule(graph=graph, rule_id="rule-1")
    assert result == {"deleted": True, "id": "rule-1"}
    inbox.message_rules.by_message_rule_id.assert_called_once_with("rule-1")
    by_id.delete.assert_awaited_once()


@pytest.mark.asyncio
async def test_delete_rule_requires_rule_id():
    from outlook_mcp.graph.errors import GraphValidationError

    graph = MagicMock()
    with pytest.raises(GraphValidationError):
        await mail_rules.delete_rule(graph=graph, rule_id="")


# Regression: kiota's BackingStoreSerializationWriterProxyFactory emits
# `write_null_value(name)` on the *parent* writer for every attribute that was
# explicitly assigned `None`. That conflicts with the serialized object dict
# and raises "Invalid Json output". The helpers must only set non-None attrs.


def test_build_predicates_skips_none_fields():
    pred = mail_rules._build_predicates(
        sender_contains=["x"],
        subject_contains=None,
        body_contains=None,
        body_or_subject_contains=None,
        from_addresses=None,
        has_attachments=None,
    )
    assert pred is not None
    # backing_store tracks explicit assignments; only sender_contains should be tracked.
    changed = pred.backing_store.enumerate_keys_for_values_changed_to_null()
    assert list(changed) == []
    assert pred.sender_contains == ["x"]


def test_build_actions_skips_none_fields():
    actions = mail_rules._build_actions(
        move_to_folder="fid",
        mark_as_read=None,
        delete=None,
        stop_processing_rules=True,
    )
    assert actions is not None
    changed = actions.backing_store.enumerate_keys_for_values_changed_to_null()
    assert list(changed) == []
    assert actions.move_to_folder == "fid"
    assert actions.stop_processing_rules is True


@pytest.mark.asyncio
async def test_create_rule_defaults_sequence_to_1():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    created = _fake_rule()
    inbox.message_rules.post = AsyncMock(return_value=created)

    await mail_rules.create_rule(
        graph=graph,
        display_name="r",
        sender_contains=["x"],
        move_to_folder="fid",
    )
    # Graph rejects sequence=0; we default to 1 so the rule is accepted.
    posted = inbox.message_rules.post.await_args.args[0]
    assert posted.sequence == 1


@pytest.mark.asyncio
async def test_create_rule_respects_explicit_sequence():
    graph, mb, inbox = _build_graph_with_rules_endpoint()
    created = _fake_rule()
    inbox.message_rules.post = AsyncMock(return_value=created)

    await mail_rules.create_rule(
        graph=graph,
        display_name="r",
        sender_contains=["x"],
        move_to_folder="fid",
        sequence=7,
    )
    posted = inbox.message_rules.post.await_args.args[0]
    assert posted.sequence == 7
