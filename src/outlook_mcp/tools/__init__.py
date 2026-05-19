"""Tool registration entry point.

register_all(mcp) wires every tool into the FastMCP app. Each tool module
exposes a register(mcp, *, graph) function so we can keep modules small.
"""

from __future__ import annotations

from outlook_mcp.tools import util


def register_all(mcp, *, graph) -> None:
    util.register(mcp, graph=graph)
