from pathlib import Path

import pytest

from outlook_mcp import config


def test_require_env_returns_value_when_set(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_CLIENT_ID", "fake-client-id")
    assert config.require_env("OUTLOOK_MCP_CLIENT_ID") == "fake-client-id"


def test_require_env_raises_when_missing(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_CLIENT_ID", raising=False)
    with pytest.raises(config.ConfigError) as exc:
        config.require_env("OUTLOOK_MCP_CLIENT_ID")
    assert "OUTLOOK_MCP_CLIENT_ID" in str(exc.value)


def test_token_cache_path_uses_default_when_unset(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", raising=False)
    path = config.token_cache_path()
    assert path == Path.home() / ".outlook-mcp" / "token_cache.bin"


def test_token_cache_path_uses_override_with_tilde_expansion(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", "~/custom/cache.bin")
    path = config.token_cache_path()
    assert path == Path.home() / "custom" / "cache.bin"


def test_scopes_match_spec():
    # Spec section 5.3
    assert config.SCOPES == [
        "Mail.ReadWrite",
        "Mail.ReadWrite.Shared",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
        "MailboxSettings.ReadWrite",
        "User.Read",
    ]


def test_authority_uses_tenant_id(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_TENANT_ID", "abc-tenant")
    assert config.authority() == "https://login.microsoftonline.com/abc-tenant"
