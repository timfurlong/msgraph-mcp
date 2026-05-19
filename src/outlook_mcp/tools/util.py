"""Utility tools (whoami)."""

from __future__ import annotations

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import map_kiota_error
from outlook_mcp.graph.serialize import user_to_dict
from outlook_mcp.graph.trimming import trim_user


async def whoami(*, graph, include_raw: bool = False) -> dict:
    """Return the signed-in user's profile.

    Also a smoke test that authentication is working.

    Args:
        include_raw: when True, include the raw Graph payload under "raw".

    Returns:
        Trimmed user object: {id, display_name, user_principal_name, mail, job_title}
    """
    try:
        user = await graph.mailbox(None).get()
    except NotAuthenticatedError:
        # Surface the login-required hint directly; don't wrap as a Graph error.
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc
    return trim_user(user_to_dict(user), include_raw=include_raw)


def register(mcp, *, graph) -> None:
    @mcp.tool(name="whoami", description=whoami.__doc__ or "")
    async def _whoami(include_raw: bool = False) -> dict:
        return await whoami(graph=graph, include_raw=include_raw)
