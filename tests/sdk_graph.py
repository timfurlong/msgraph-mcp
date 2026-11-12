"""A GraphClient backed by the real msgraph SDK over a mock HTTP transport.

Tools built on it send real requests through the SDK's middleware, so tests
can check the exact URL a tool sends and answer with real Graph response
bodies, including error bodies. No auth is involved.
"""

from __future__ import annotations

from collections.abc import Callable

import httpx
from kiota_abstractions.authentication import AnonymousAuthenticationProvider
from msgraph.graph_request_adapter import GraphRequestAdapter, options
from msgraph.graph_service_client import GraphServiceClient
from msgraph_core import GraphClientFactory

from msgraph_mcp.graph.client import GraphClient

Handler = Callable[[httpx.Request], httpx.Response]


def sdk_graph(handler: Handler) -> GraphClient:
    http = GraphClientFactory.create_with_default_middleware(
        options=options,
        client=httpx.AsyncClient(transport=httpx.MockTransport(handler)),
    )
    sdk = GraphServiceClient(
        request_adapter=GraphRequestAdapter(
            AnonymousAuthenticationProvider(), client=http
        )
    )
    return GraphClient(sdk=sdk)


def graph_error(status: int, code: str, message: str) -> httpx.Response:
    """A Graph error response, shaped the way Graph returns one."""
    return httpx.Response(status, json={"error": {"code": code, "message": message}})
