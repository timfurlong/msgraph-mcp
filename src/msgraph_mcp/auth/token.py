"""Silent token acquisition for the running MCP server.

The server NEVER runs device code flow. If silent acquisition fails,
get_access_token raises NotAuthenticatedError pointing at the login CLI.
"""

from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError

from msgraph_mcp import config
from msgraph_mcp.auth.msal_app import build_app, persist_cache


class NotAuthenticatedError(ToolError, RuntimeError):
    """Raised when no valid token is available and silent refresh failed.

    A ToolError so MCPServer passes the login hint through to the caller.
    """


_LOGIN_HINT = "Not authenticated. Run `msgraph-mcp-login` in a terminal to sign in."


def get_access_token() -> str:
    app, cache = build_app()
    accounts = app.get_accounts()
    if not accounts:
        raise NotAuthenticatedError(_LOGIN_HINT)
    result = app.acquire_token_silent(config.SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise NotAuthenticatedError(_LOGIN_HINT)
    if cache.has_state_changed:
        persist_cache(cache)
    return result["access_token"]
