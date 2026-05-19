from unittest.mock import MagicMock, patch

from outlook_mcp.graph import client as client_mod


def _make_client_with_fakes():
    fake_sdk = MagicMock()
    fake_sdk.me = MagicMock(name="me_builder")
    fake_sdk.users.by_user_id = MagicMock(return_value=MagicMock(name="users_builder"))
    with patch("outlook_mcp.graph.client.GraphServiceClient", return_value=fake_sdk):
        c = client_mod.GraphClient()
    return c, fake_sdk


def test_mailbox_routes_to_me_when_none():
    c, fake_sdk = _make_client_with_fakes()
    assert c.mailbox(None) is fake_sdk.me


def test_mailbox_routes_to_user_when_given():
    c, fake_sdk = _make_client_with_fakes()
    result = c.mailbox("alice@example.com")
    fake_sdk.users.by_user_id.assert_called_once_with("alice@example.com")
    assert result is fake_sdk.users.by_user_id.return_value


def test_raw_returns_underlying_sdk_client():
    c, fake_sdk = _make_client_with_fakes()
    assert c.raw is fake_sdk
