"""Graph error mapping.

Kiota / msgraph-sdk raises a variety of exception shapes for HTTP errors.
We normalize them into a single GraphAPIError class that the tool layer
can catch and turn into a clean MCP tool error.
"""

from __future__ import annotations

from typing import Any


class GraphValidationError(RuntimeError):
    """Raised when tool params fail validation before the Graph call."""


class GraphAPIError(RuntimeError):
    """Raised when Graph returns a non-success response."""

    def __init__(self, *, status: int, code: str, message: str) -> None:
        self.status = status
        self.code = code
        self.message = message
        super().__init__(f"Graph API {status}: {code} — {message}")


def map_kiota_error(exc: Any) -> GraphAPIError:
    """Map an arbitrary kiota/msgraph exception into a GraphAPIError.

    Kiota's ODataError exposes:
      - response_status_code: int
      - error: object with .code (str) and .message (str)
    Some lower-level errors only expose a status and a string. For anything
    we can't introspect, fall back to status=0 and stringify the exception.
    """
    status = int(getattr(exc, "response_status_code", 0) or 0)
    inner = getattr(exc, "error", None)
    code = "Unknown"
    message = str(exc) if status == 0 else f"HTTP {status}"
    if inner is not None:
        code = getattr(inner, "code", code) or code
        message = getattr(inner, "message", message) or message
    return GraphAPIError(status=status, code=code, message=message)
