"""Configuration loading for outlook_mcp.

Env vars are read once at import-side via require_env. Values are sourced from
the process environment with .env as a fallback (process wins via override=False).
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv


# Load .env once at module import; process env takes precedence.
load_dotenv(override=False)


SCOPES: list[str] = [
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "User.Read",
]


DEFAULT_CACHE_PATH = Path.home() / ".outlook-mcp" / "token_cache.bin"


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or invalid."""


def require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise ConfigError(
            f"Missing required env var: {name}. "
            f"Set it in the process env or in a .env file at the repo root."
        )
    return value


def token_cache_path() -> Path:
    override = os.environ.get("OUTLOOK_MCP_TOKEN_CACHE_PATH")
    if override:
        return Path(override).expanduser()
    return DEFAULT_CACHE_PATH


def authority() -> str:
    tenant = require_env("OUTLOOK_MCP_TENANT_ID")
    return f"https://login.microsoftonline.com/{tenant}"


def client_id() -> str:
    return require_env("OUTLOOK_MCP_CLIENT_ID")
