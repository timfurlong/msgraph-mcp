"""Configuration loading for msgraph_mcp.

Env vars are read once at import-side via require_env. Values are sourced from
the process environment with .env as a fallback (process wins via override=False).
.env is looked for in the working directory first, then relative to this module.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import find_dotenv, load_dotenv


# Load .env once at module import; process env takes precedence.
#
# The working-directory search runs first because that is where a user expects
# their .env to be read from. Bare load_dotenv() resolves relative to this
# module instead, which lands inside site-packages for an installed copy and
# never finds it; it stays as a second pass so a source checkout launched from
# an unrelated directory keeps working. override=False throughout, so the
# process env wins over both and the first file found wins over the second.
load_dotenv(find_dotenv(usecwd=True), override=False)
load_dotenv(override=False)


SCOPES: list[str] = [
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "MailboxSettings.ReadWrite",
    "User.Read",
    # Teams read-history (ChannelMessage.Read.All requires tenant admin consent)
    "Chat.Read",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
]


DEFAULT_CACHE_PATH = Path.home() / ".msgraph-mcp" / "token_cache.bin"

# Pre-rename (outlook-mcp) locations, honored so existing setups keep working.
LEGACY_CACHE_PATH = Path.home() / ".outlook-mcp" / "token_cache.bin"
_LEGACY_ENV_PREFIX = "OUTLOOK_MCP_"


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or invalid."""


def _env(name: str) -> str | None:
    """Read an MSGRAPH_MCP_* env var, falling back to its legacy OUTLOOK_MCP_* name."""
    value = os.environ.get(name)
    if value:
        return value
    return os.environ.get(name.replace("MSGRAPH_MCP_", _LEGACY_ENV_PREFIX, 1))


def require_env(name: str) -> str:
    value = _env(name)
    if not value:
        raise ConfigError(
            f"Missing required env var: {name}. "
            f"Set it in the process env or in a .env file at the repo root."
        )
    return value


def token_cache_path() -> Path:
    override = _env("MSGRAPH_MCP_TOKEN_CACHE_PATH")
    if override:
        return Path(override).expanduser()
    if not DEFAULT_CACHE_PATH.exists() and LEGACY_CACHE_PATH.exists():
        return LEGACY_CACHE_PATH
    return DEFAULT_CACHE_PATH


def authority() -> str:
    tenant = require_env("MSGRAPH_MCP_TENANT_ID")
    return f"https://login.microsoftonline.com/{tenant}"


def client_id() -> str:
    return require_env("MSGRAPH_MCP_CLIENT_ID")
