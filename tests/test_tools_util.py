from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphAPIError
from outlook_mcp.tools import util as util_tools


@pytest.mark.asyncio
async def test_whoami_returns_trimmed_user():
    fake_graph = MagicMock()
    me_builder = MagicMock()
    user_obj = MagicMock(
        additional_data={},
        id="u1",
        display_name="Alice",
        user_principal_name="alice@example.com",
        mail="alice@example.com",
        job_title="Engineer",
    )
    me_builder.get = AsyncMock(return_value=user_obj)
    fake_graph.mailbox = MagicMock(return_value=me_builder)

    result = await util_tools.whoami(graph=fake_graph)

    fake_graph.mailbox.assert_called_once_with(None)
    assert result["display_name"] == "Alice"
    assert result["mail"] == "alice@example.com"
    assert "raw" not in result


@pytest.mark.asyncio
async def test_whoami_with_include_raw():
    fake_graph = MagicMock()
    me_builder = MagicMock()
    user_obj = MagicMock(
        additional_data={"extra": "x"},
        id="u1",
        display_name="Alice",
        user_principal_name="alice@example.com",
        mail="alice@example.com",
        job_title="Engineer",
    )
    me_builder.get = AsyncMock(return_value=user_obj)
    fake_graph.mailbox = MagicMock(return_value=me_builder)

    result = await util_tools.whoami(graph=fake_graph, include_raw=True)
    assert "raw" in result


@pytest.mark.asyncio
async def test_whoami_propagates_not_authenticated_error():
    fake_graph = MagicMock()
    me_builder = MagicMock()
    me_builder.get = AsyncMock(side_effect=NotAuthenticatedError("login required"))
    fake_graph.mailbox = MagicMock(return_value=me_builder)

    with pytest.raises(NotAuthenticatedError):
        await util_tools.whoami(graph=fake_graph)


@pytest.mark.asyncio
async def test_whoami_maps_other_exceptions_to_graph_api_error():
    fake_graph = MagicMock()
    me_builder = MagicMock()
    me_builder.get = AsyncMock(side_effect=ValueError("boom"))
    fake_graph.mailbox = MagicMock(return_value=me_builder)

    with pytest.raises(GraphAPIError) as exc:
        await util_tools.whoami(graph=fake_graph)
    assert "boom" in str(exc.value)
