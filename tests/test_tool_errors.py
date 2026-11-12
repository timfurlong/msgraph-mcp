"""Tool failures reach the MCP caller with their details.

MCPServer reports any exception that isn't a ToolError as a bare
"Error executing tool <name>" and keeps its text on the server. These tests
call tools through a real in-process MCP client, so they check exactly what
the caller sees.
"""

from unittest.mock import MagicMock, patch

import httpx
import pytest
from mcp import Client
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult, TextContent

from msgraph_mcp import server
from msgraph_mcp.auth.token import NotAuthenticatedError
from msgraph_mcp.graph.batch import GraphBatchTransport
from msgraph_mcp.graph.errors import GraphAPIError, GraphValidationError
from tests.sdk_graph import graph_error, sdk_graph


@pytest.fixture
def fake_env(monkeypatch):
    monkeypatch.setenv("MSGRAPH_MCP_CLIENT_ID", "fake-client-id")
    monkeypatch.setenv("MSGRAPH_MCP_TENANT_ID", "fake-tenant-id")


async def _call(graph, tool: str, arguments: dict) -> CallToolResult:
    with patch("msgraph_mcp.server.GraphClient", return_value=graph):
        mcp = server.build_server()
    async with Client(mcp) as client:
        return await client.call_tool(tool, arguments)


def _text(result: CallToolResult) -> str:
    return " ".join(c.text for c in result.content if isinstance(c, TextContent))


@pytest.mark.parametrize(
    "error_class", [GraphAPIError, GraphValidationError, NotAuthenticatedError]
)
def test_anticipated_errors_are_tool_errors(error_class):
    assert issubclass(error_class, ToolError)


@pytest.mark.parametrize(
    ("tool", "arguments"),
    [
        (
            "list_messages",
            {"folder_id": "sentitems", "filter": "contains(subject,'x')"},
        ),
        ("search_messages", {"query": "report"}),
        ("get_message", {"message_id": "m1"}),
        ("list_folders", {}),
        ("list_rules", {}),
        (
            "list_events",
            {
                "calendar_id": "c1",
                "start_datetime": "2026-11-01T00:00:00Z",
                "end_datetime": "2026-11-08T00:00:00Z",
            },
        ),
        ("list_chats", {}),
        ("whoami", {}),
        ("batch_mark_read", {"message_ids": ["m1", "m2"]}),
    ],
)
async def test_graph_error_reaches_caller(fake_env, tool, arguments):
    graph = sdk_graph(
        lambda request: graph_error(
            400,
            "InefficientFilter",
            "The restriction or sort order is too complex for this operation.",
        )
    )

    result = await _call(graph, tool, arguments)

    assert result.is_error
    text = _text(result)
    assert "400" in text
    assert "InefficientFilter" in text
    assert "The restriction or sort order is too complex" in text


async def test_validation_error_reaches_caller(fake_env):
    result = await _call(MagicMock(), "list_messages", {"limit": 200})

    assert result.is_error
    assert "limit must be an integer in 1..100" in _text(result)


async def test_login_hint_reaches_caller(fake_env):
    graph = MagicMock()
    graph.mailbox.return_value.get.side_effect = NotAuthenticatedError(
        "Not authenticated. Run `msgraph-mcp-login` in a terminal to sign in."
    )

    result = await _call(graph, "whoami", {})

    assert result.is_error
    assert "msgraph-mcp-login" in _text(result)


async def test_batch_transport_maps_graph_errors():
    graph = sdk_graph(
        lambda request: graph_error(
            401, "InvalidAuthenticationToken", "Access token has expired."
        )
    )

    with pytest.raises(GraphAPIError) as excinfo:
        await GraphBatchTransport(graph).post_batch({"requests": []})

    assert excinfo.value.status == 401
    assert excinfo.value.code == "InvalidAuthenticationToken"
    assert excinfo.value.message == "Access token has expired."


async def test_batch_transport_passes_through_success():
    graph = sdk_graph(
        lambda request: httpx.Response(200, json={"responses": [{"id": "1"}]})
    )

    out = await GraphBatchTransport(graph).post_batch({"requests": []})

    assert out == {"responses": [{"id": "1"}]}
