import stat
from unittest.mock import patch

import pytest

from outlook_mcp.auth import msal_app


@pytest.fixture
def fake_env(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_CLIENT_ID", "fake-client-id")
    monkeypatch.setenv("OUTLOOK_MCP_TENANT_ID", "fake-tenant-id")


@pytest.fixture
def patched_pca():
    # MSAL's PublicClientApplication performs network discovery on
    # construction; patch it out so we can exercise build_app's own logic
    # (cache loading, authority URL wiring) without network.
    with patch.object(msal_app.msal, "PublicClientApplication") as mock_pca:
        yield mock_pca


def test_build_app_passes_tenant_authority_to_msal(fake_env, tmp_path, monkeypatch, patched_pca):
    monkeypatch.setenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", str(tmp_path / "cache.bin"))
    msal_app.build_app()
    kwargs = patched_pca.call_args.kwargs
    assert kwargs["authority"] == "https://login.microsoftonline.com/fake-tenant-id"
    assert kwargs["client_id"] == "fake-client-id"


def test_build_app_loads_existing_cache(fake_env, tmp_path, monkeypatch, patched_pca):
    cache_path = tmp_path / "cache.bin"
    # MSAL cache files start with a JSON object; the empty {} is valid.
    cache_path.write_text("{}")
    monkeypatch.setenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", str(cache_path))
    _app, cache = msal_app.build_app()
    # Calling serialize should round-trip without error
    assert cache.serialize().startswith("{")


def test_build_app_tolerates_missing_cache(fake_env, tmp_path, monkeypatch, patched_pca):
    monkeypatch.setenv(
        "OUTLOOK_MCP_TOKEN_CACHE_PATH", str(tmp_path / "does-not-exist.bin")
    )
    _app, cache = msal_app.build_app()
    # Fresh cache should still serialize.
    assert cache.serialize() == "{}"


def test_persist_cache_writes_0600(tmp_path, fake_env, monkeypatch, patched_pca):
    cache_path = tmp_path / "subdir" / "cache.bin"
    monkeypatch.setenv("OUTLOOK_MCP_TOKEN_CACHE_PATH", str(cache_path))
    _app, cache = msal_app.build_app()
    # Force a state change by serializing into a fake; then persist.
    msal_app.persist_cache(cache)
    assert cache_path.exists()
    mode = stat.S_IMODE(cache_path.stat().st_mode)
    assert mode == 0o600
    # Parent dir should be 0700
    parent_mode = stat.S_IMODE(cache_path.parent.stat().st_mode)
    assert parent_mode == 0o700
