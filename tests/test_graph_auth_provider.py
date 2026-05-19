import time
from unittest.mock import patch

from outlook_mcp.graph import auth_provider


def test_credential_get_token_returns_access_token():
    with patch("outlook_mcp.graph.auth_provider.get_access_token", return_value="tok-xyz"):
        cred = auth_provider.MsalTokenCredential()
        result = cred.get_token("https://graph.microsoft.com/.default")
    assert result.token == "tok-xyz"
    # expires_on should be in the future
    assert result.expires_on > int(time.time())


def test_credential_ignores_scope_arg():
    with patch("outlook_mcp.graph.auth_provider.get_access_token", return_value="tok-xyz") as p:
        cred = auth_provider.MsalTokenCredential()
        cred.get_token("Mail.ReadWrite", "Calendars.ReadWrite")
    # get_access_token should be called without args (it uses MSAL's pre-configured scopes)
    p.assert_called_once_with()
