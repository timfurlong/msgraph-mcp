"""Bridge our MSAL silent-acquire flow into msgraph-sdk via TokenCredential.

msgraph-sdk's GraphServiceClient takes an azure-core TokenCredential and
calls .get_token(*scopes) on every authenticated request. We ignore the
scopes argument — our MSAL PublicClientApplication is pre-configured with
the right delegated scopes during login, and the cached token already has
the correct audience.
"""

from __future__ import annotations

import time

from azure.core.credentials import AccessToken, TokenCredential

from outlook_mcp.auth.token import get_access_token


# Short expiry: MSAL's cache handles real refresh. Each get_token call goes
# through get_access_token() which calls acquire_token_silent (cheap, cached).
_BUFFER_SECONDS = 300


class MsalTokenCredential(TokenCredential):
    """TokenCredential that delegates to our MSAL silent-acquire pipeline."""

    def get_token(self, *scopes: str, **kwargs) -> AccessToken:  # type: ignore[override]
        token = get_access_token()
        return AccessToken(token, int(time.time() + _BUFFER_SECONDS))
