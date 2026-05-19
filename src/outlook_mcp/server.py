"""FastMCP server entry point.

Boot order:
  1. Validate config (fail fast if env is missing).
  2. Build the GraphClient (lazy token acquisition — no Graph calls yet).
  3. Build a FastMCP app, register all tools, and run on stdio.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_mcp import config
from outlook_mcp.graph.client import GraphClient
from outlook_mcp.tools import register_all


_INSTRUCTIONS = """
Microsoft Outlook (mail + calendar) via Microsoft Graph.

This server acts as the signed-in user (delegated auth). All tools that
touch a mailbox or calendar accept an optional `mailbox` argument (email
or user ID) to target shared mailboxes/calendars; omit it to use the
signed-in user's own mailbox.

Pagination: list/search tools accept `limit` (1-100, default 25) and
`page_token`. Pass the returned `next_page_token` back as `page_token`
to fetch the next page.

Responses are trimmed by default. Pass `include_raw=true` to any tool
that returns objects to also get the full Graph payload.
""".strip()


def build_server() -> FastMCP:
    # Fail fast on missing env (raises ConfigError before we hit the event loop)
    config.client_id()
    config.authority()

    mcp = FastMCP(name="outlook-mcp", instructions=_INSTRUCTIONS)
    graph = GraphClient()
    register_all(mcp, graph=graph)
    return mcp


def main() -> None:
    mcp = build_server()
    mcp.run()  # stdio transport by default
