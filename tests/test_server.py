from unittest.mock import patch

import pytest

from msgraph_mcp import server


@pytest.fixture
def fake_env(monkeypatch):
    monkeypatch.setenv("MSGRAPH_MCP_CLIENT_ID", "fake-client-id")
    monkeypatch.setenv("MSGRAPH_MCP_TENANT_ID", "fake-tenant-id")


async def test_build_server_registers_whoami(fake_env):
    with patch("msgraph_mcp.server.GraphClient") as fake_graph_class:
        mcp = server.build_server()
    # GraphClient should have been constructed once during build.
    fake_graph_class.assert_called_once()
    names = {tool.name for tool in await mcp.list_tools()}
    assert "whoami" in names


def test_main_calls_run(monkeypatch, fake_env):
    calls: list = []

    def fake_run(self, *args, **kwargs):
        calls.append((args, kwargs))

    monkeypatch.setattr("mcp.server.mcpserver.MCPServer.run", fake_run)
    with patch("msgraph_mcp.server.GraphClient"):
        server.main()
    assert len(calls) == 1
