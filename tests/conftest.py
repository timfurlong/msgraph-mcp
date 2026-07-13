"""Pytest config: gate live integration tests on MSGRAPH_MCP_INTEGRATION=1."""

from __future__ import annotations

import os

import pytest


def pytest_collection_modifyitems(config, items):
    if os.environ.get("MSGRAPH_MCP_INTEGRATION") == "1":
        return
    skip_marker = pytest.mark.skip(
        reason="set MSGRAPH_MCP_INTEGRATION=1 to run live integration tests"
    )
    for item in items:
        if "tests/integration/" in str(item.fspath):
            item.add_marker(skip_marker)
