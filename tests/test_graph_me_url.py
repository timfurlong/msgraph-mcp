"""Guard the /me URL rewrite, which lives entirely inside the SDK stack.

GraphServiceClient.me builds /users/me-token-to-replace and relies on
msgraph-core's middleware pipeline to rewrite it to /me. kiota-http 1.13 broke
that pipeline for msgraph-core <= 1.5.1 (see the bound in pyproject.toml), and
every call failed against live Graph while the mocked unit tests stayed green.
This runs the real SDK pipeline over a mock transport, so it needs no auth.
"""

import httpx
from kiota_abstractions.authentication import AnonymousAuthenticationProvider
from msgraph.graph_request_adapter import GraphRequestAdapter, options
from msgraph.graph_service_client import GraphServiceClient
from msgraph_core import GraphClientFactory


async def test_me_is_rewritten_before_it_reaches_the_wire():
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(str(request.url))
        return httpx.Response(200, json={"id": "user-1"})

    http = GraphClientFactory.create_with_default_middleware(
        options=options,
        client=httpx.AsyncClient(transport=httpx.MockTransport(handler)),
    )
    sdk = GraphServiceClient(
        request_adapter=GraphRequestAdapter(
            AnonymousAuthenticationProvider(), client=http
        )
    )

    await sdk.me.get()

    assert seen == ["https://graph.microsoft.com/v1.0/me"]
