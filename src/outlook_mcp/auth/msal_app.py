"""MSAL PublicClientApplication builder + token cache persistence."""

from __future__ import annotations

import os

import msal

from outlook_mcp import config


def build_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    """Build the MSAL PublicClientApplication and bind a SerializableTokenCache.

    Returns the (app, cache) tuple. The cache is loaded from disk if a cache
    file exists at the configured path; otherwise it starts empty.

    Call persist_cache(cache) after any operation that may have mutated it
    (acquire_token_silent rotates refresh tokens; the cache reports
    has_state_changed=True in that case).
    """
    cache = msal.SerializableTokenCache()
    cache_path = config.token_cache_path()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text())
    app = msal.PublicClientApplication(
        client_id=config.client_id(),
        authority=config.authority(),
        token_cache=cache,
    )
    return app, cache


def persist_cache(cache: msal.SerializableTokenCache) -> None:
    """Write the cache to disk with 0600 permissions, creating parent dirs at 0700.

    Safe to call even if has_state_changed is False; we always write to keep
    behavior simple and predictable. Caller may guard with `if cache.has_state_changed`
    if write frequency matters.
    """
    cache_path = config.token_cache_path()
    parent = cache_path.parent
    parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    # Ensure parent dir perms even if it already existed
    os.chmod(parent, 0o700)
    cache_path.write_text(cache.serialize())
    os.chmod(cache_path, 0o600)
