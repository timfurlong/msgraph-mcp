"""Small wrapper around msgraph-sdk's GraphServiceClient.

Centralizes:
  - GraphServiceClient construction with our MsalTokenCredential
  - /me vs /users/{id} routing for shared mailboxes/calendars
  - A handle to the raw SDK client for cases that need it directly
"""

from __future__ import annotations

from msgraph.graph_service_client import GraphServiceClient

from outlook_mcp import config
from outlook_mcp.graph.auth_provider import MsalTokenCredential


class GraphClient:
    def __init__(self) -> None:
        self._sdk = GraphServiceClient(
            credentials=MsalTokenCredential(),
            scopes=config.SCOPES,
        )

    @property
    def raw(self) -> GraphServiceClient:
        return self._sdk

    def mailbox(self, mailbox: str | None):
        """Return the builder rooted at /me or /users/{mailbox}.

        Use it like:
            await graph.mailbox(mailbox).messages.get(...)
            await graph.mailbox(mailbox).calendar.events.get(...)
        """
        if mailbox is None:
            return self._sdk.me
        return self._sdk.users.by_user_id(mailbox)
