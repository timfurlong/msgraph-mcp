from pathlib import Path

import pytest

from msgraph_mcp import config


def test_require_env_returns_value_when_set(monkeypatch):
    monkeypatch.setenv("MSGRAPH_MCP_CLIENT_ID", "fake-client-id")
    assert config.require_env("MSGRAPH_MCP_CLIENT_ID") == "fake-client-id"


def test_require_env_falls_back_to_legacy_name(monkeypatch):
    monkeypatch.delenv("MSGRAPH_MCP_CLIENT_ID", raising=False)
    monkeypatch.setenv("OUTLOOK_MCP_CLIENT_ID", "legacy-client-id")
    assert config.require_env("MSGRAPH_MCP_CLIENT_ID") == "legacy-client-id"


def test_require_env_raises_when_missing(monkeypatch):
    monkeypatch.delenv("MSGRAPH_MCP_CLIENT_ID", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_CLIENT_ID", raising=False)
    with pytest.raises(config.ConfigError) as exc:
        config.require_env("MSGRAPH_MCP_CLIENT_ID")
    assert "MSGRAPH_MCP_CLIENT_ID" in str(exc.value)


def test_token_cache_path_uses_default_when_unset(monkeypatch, tmp_path):
    monkeypatch.delenv("MSGRAPH_MCP_TOKEN_CACHE_PATH", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", raising=False)
    monkeypatch.setattr(config, "DEFAULT_CACHE_PATH", tmp_path / "new" / "token_cache.bin")
    monkeypatch.setattr(config, "LEGACY_CACHE_PATH", tmp_path / "old" / "token_cache.bin")
    assert config.token_cache_path() == tmp_path / "new" / "token_cache.bin"


def test_token_cache_path_falls_back_to_legacy_cache(monkeypatch, tmp_path):
    monkeypatch.delenv("MSGRAPH_MCP_TOKEN_CACHE_PATH", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", raising=False)
    legacy = tmp_path / "old" / "token_cache.bin"
    legacy.parent.mkdir()
    legacy.write_bytes(b"cache")
    monkeypatch.setattr(config, "DEFAULT_CACHE_PATH", tmp_path / "new" / "token_cache.bin")
    monkeypatch.setattr(config, "LEGACY_CACHE_PATH", legacy)
    assert config.token_cache_path() == legacy


def test_token_cache_path_prefers_new_default_over_legacy(monkeypatch, tmp_path):
    monkeypatch.delenv("MSGRAPH_MCP_TOKEN_CACHE_PATH", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", raising=False)
    new = tmp_path / "new" / "token_cache.bin"
    legacy = tmp_path / "old" / "token_cache.bin"
    for p in (new, legacy):
        p.parent.mkdir()
        p.write_bytes(b"cache")
    monkeypatch.setattr(config, "DEFAULT_CACHE_PATH", new)
    monkeypatch.setattr(config, "LEGACY_CACHE_PATH", legacy)
    assert config.token_cache_path() == new


def test_token_cache_path_uses_override_with_tilde_expansion(monkeypatch):
    monkeypatch.setenv("MSGRAPH_MCP_TOKEN_CACHE_PATH", "~/custom/cache.bin")
    path = config.token_cache_path()
    assert path == Path.home() / "custom" / "cache.bin"


def test_scopes_match_spec():
    # Mail/calendar scopes (outlook design 5.3) + Teams read scopes (teams design)
    assert config.SCOPES == [
        "Mail.ReadWrite",
        "Mail.ReadWrite.Shared",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
        "MailboxSettings.ReadWrite",
        "User.Read",
        "Chat.Read",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
    ]


def test_scopes_include_teams_read():
    from msgraph_mcp import config
    for scope in ("Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All"):
        assert scope in config.SCOPES


def test_authority_uses_tenant_id(monkeypatch):
    monkeypatch.setenv("MSGRAPH_MCP_TENANT_ID", "abc-tenant")
    assert config.authority() == "https://login.microsoftonline.com/abc-tenant"
