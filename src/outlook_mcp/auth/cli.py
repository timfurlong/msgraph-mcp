"""outlook-mcp-login entry point: runs the device code flow interactively."""

from __future__ import annotations

import sys

from outlook_mcp import config
from outlook_mcp.auth.msal_app import build_app, persist_cache


def run_login() -> int:
    app, cache = build_app()
    flow = app.initiate_device_flow(scopes=config.SCOPES)
    if "user_code" not in flow:
        print(
            f"Failed to start device flow: {flow.get('error', flow)}",
            file=sys.stderr,
        )
        return 1

    # Message contains the verification URL and the code; print to stdout
    # so the user sees it in their terminal.
    print(flow["message"], flush=True)

    result = app.acquire_token_by_device_flow(flow)  # blocks until user completes flow
    if "error" in result:
        print(
            f"Login failed: {result['error']} — {result.get('error_description', '')}",
            file=sys.stderr,
        )
        return 1

    persist_cache(cache)

    username = (
        result.get("id_token_claims", {}).get("preferred_username")
        or "<unknown>"
    )
    print(f"Logged in as {username}", flush=True)
    print(f"Token cache saved to {config.token_cache_path()}", flush=True)
    return 0


def main() -> None:
    sys.exit(run_login())
