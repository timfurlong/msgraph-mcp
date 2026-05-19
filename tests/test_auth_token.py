from unittest.mock import MagicMock

import pytest

from outlook_mcp.auth import token as token_mod


def test_get_access_token_returns_token_when_silent_succeeds(monkeypatch):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_cache.has_state_changed = False
    fake_app.get_accounts.return_value = [{"username": "me@example.com"}]
    fake_app.acquire_token_silent.return_value = {"access_token": "tok-123"}

    monkeypatch.setattr(token_mod, "build_app", lambda: (fake_app, fake_cache))

    assert token_mod.get_access_token() == "tok-123"


def test_get_access_token_raises_when_no_accounts(monkeypatch):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_app.get_accounts.return_value = []
    monkeypatch.setattr(token_mod, "build_app", lambda: (fake_app, fake_cache))

    with pytest.raises(token_mod.NotAuthenticatedError) as exc:
        token_mod.get_access_token()
    assert "outlook-mcp-login" in str(exc.value)


def test_get_access_token_raises_when_silent_returns_none(monkeypatch):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_app.get_accounts.return_value = [{"username": "me@example.com"}]
    fake_app.acquire_token_silent.return_value = None
    monkeypatch.setattr(token_mod, "build_app", lambda: (fake_app, fake_cache))

    with pytest.raises(token_mod.NotAuthenticatedError):
        token_mod.get_access_token()


def test_get_access_token_raises_when_silent_returns_error(monkeypatch):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_app.get_accounts.return_value = [{"username": "me@example.com"}]
    fake_app.acquire_token_silent.return_value = {"error": "invalid_grant"}
    monkeypatch.setattr(token_mod, "build_app", lambda: (fake_app, fake_cache))

    with pytest.raises(token_mod.NotAuthenticatedError):
        token_mod.get_access_token()


def test_get_access_token_persists_cache_when_state_changed(monkeypatch):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_cache.has_state_changed = True
    fake_app.get_accounts.return_value = [{"username": "me@example.com"}]
    fake_app.acquire_token_silent.return_value = {"access_token": "tok-123"}

    persist_calls: list = []
    monkeypatch.setattr(token_mod, "build_app", lambda: (fake_app, fake_cache))
    monkeypatch.setattr(token_mod, "persist_cache", lambda c: persist_calls.append(c))

    token_mod.get_access_token()
    assert persist_calls == [fake_cache]
