"""Opaque pagination tokens and limit validation.

The agent never sees raw @odata.nextLink URLs. We base64-encode the URL
and validate on decode that it points at graph.microsoft.com.
"""

from __future__ import annotations

import base64
import binascii
from urllib.parse import urlparse

from outlook_mcp.graph.errors import GraphValidationError


_ALLOWED_HOSTS = {"graph.microsoft.com"}


def encode_next_link(url: str | None) -> str | None:
    if not url:
        return None
    return base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")


def decode_page_token(token: str) -> str:
    try:
        # Re-pad
        padded = token + "=" * (-len(token) % 4)
        raw = base64.urlsafe_b64decode(padded.encode("ascii"))
        url = raw.decode("utf-8")
    except (binascii.Error, UnicodeDecodeError) as exc:
        raise GraphValidationError(f"Invalid page_token: {exc}") from exc

    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in _ALLOWED_HOSTS:
        raise GraphValidationError(
            f"Invalid page_token: must point at graph.microsoft.com (got {parsed.hostname!r})"
        )
    return url


def validate_limit(limit: int, *, maximum: int = 100) -> int:
    if not isinstance(limit, int) or limit < 1 or limit > maximum:
        raise GraphValidationError(
            f"limit must be an integer in 1..{maximum} (got {limit!r})"
        )
    return limit
