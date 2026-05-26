"""Live end-to-end: create folder + rule, list, then clean up.

Skipped unless OUTLOOK_MCP_INTEGRATION=1 (see tests/conftest.py).
Requires a configured signed-in user (run `uv run outlook-mcp-login` first).
"""

from __future__ import annotations

import uuid

import pytest

from outlook_mcp.graph.client import GraphClient
from outlook_mcp.tools import mail_folders, mail_rules


@pytest.fixture
def graph():
    # See note in tests/integration/test_live_smoke.py: function scope avoids
    # the kiota httpx client outliving its event loop.
    return GraphClient()


@pytest.mark.asyncio
async def test_create_folder_then_rule_then_cleanup(graph):
    folder_name = f"mcp-test-{uuid.uuid4().hex[:8]}"
    rule_name = f"mcp-rule-{uuid.uuid4().hex[:8]}"

    folder = await mail_folders.create_folder(graph=graph, display_name=folder_name)
    folder_id = folder["id"]
    assert folder["display_name"] == folder_name
    rule = None
    try:
        rule = await mail_rules.create_rule(
            graph=graph,
            display_name=rule_name,
            sender_contains=["mcp-test-not-a-real-sender.invalid"],
            move_to_folder=folder_id,
            stop_processing_rules=True,
        )
        assert rule["id"]

        listed = await mail_rules.list_rules(graph=graph)
        assert any(r["id"] == rule["id"] for r in listed["items"])
    finally:
        if rule is not None:
            await mail_rules.delete_rule(graph=graph, rule_id=rule["id"])
        # delete_folder refuses non-empty folders; the folder we created has
        # no messages and no children, so it should always delete cleanly.
        await mail_folders.delete_folder(graph=graph, folder_id=folder_id)


@pytest.mark.asyncio
async def test_delete_folder_refuses_non_empty(graph):
    """End-to-end: delete_folder refuses to nuke a folder with a child folder.

    Create a parent folder + a child folder, attempt to delete the parent,
    expect GraphValidationError, then clean up child + parent in order.
    """
    from outlook_mcp.graph.errors import GraphValidationError

    parent_name = f"mcp-test-parent-{uuid.uuid4().hex[:8]}"
    child_name = f"mcp-test-child-{uuid.uuid4().hex[:8]}"
    parent = await mail_folders.create_folder(graph=graph, display_name=parent_name)
    parent_id = parent["id"]
    child = await mail_folders.create_folder(
        graph=graph, display_name=child_name, parent_folder_id=parent_id
    )
    child_id = child["id"]
    try:
        with pytest.raises(GraphValidationError, match="non-empty folder"):
            await mail_folders.delete_folder(graph=graph, folder_id=parent_id)
    finally:
        await mail_folders.delete_folder(graph=graph, folder_id=child_id)
        await mail_folders.delete_folder(graph=graph, folder_id=parent_id)
