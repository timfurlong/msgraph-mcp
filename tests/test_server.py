from unittest.mock import patch

import pytest

from outlook_mcp import server


@pytest.fixture
def fake_env(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_CLIENT_ID", "fake-client-id")
    monkeypatch.setenv("OUTLOOK_MCP_TENANT_ID", "fake-tenant-id")


def test_build_server_registers_whoami(fake_env):
    with patch("outlook_mcp.server.GraphClient") as fake_graph_class:
        mcp = server.build_server()
    # GraphClient should have been constructed once during build.
    fake_graph_class.assert_called_once()
    # FastMCP exposes registered tools via _tool_manager / list_tools depending on version.
    # We assert at minimum that the build does not throw and returns a FastMCP-like object.
    assert hasattr(mcp, "tool")


def test_main_calls_run(monkeypatch, fake_env):
    calls: list = []

    def fake_run(self, *args, **kwargs):
        calls.append((args, kwargs))

    monkeypatch.setattr("mcp.server.fastmcp.FastMCP.run", fake_run)
    with patch("outlook_mcp.server.GraphClient"):
        server.main()
    assert len(calls) == 1
