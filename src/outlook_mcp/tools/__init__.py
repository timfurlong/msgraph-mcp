"""Tool registration entry point.

register_all(mcp) wires every tool into the FastMCP app. Each tool module
exposes a register(mcp, *, graph) function so we can keep modules small.
"""

from __future__ import annotations

from outlook_mcp.tools import (
    calendar,
    mail_actions,
    mail_batch,
    mail_folders,
    mail_read,
    mail_rules,
    mail_write,
    util,
)


def register_all(mcp, *, graph) -> None:
    util.register(mcp, graph=graph)
    mail_read.register(mcp, graph=graph)
    mail_write.register(mcp, graph=graph)
    mail_folders.register(mcp, graph=graph)
    mail_actions.register(mcp, graph=graph)
    mail_batch.register(mcp, graph=graph)
    mail_rules.register(mcp, graph=graph)
    calendar.register(mcp, graph=graph)
