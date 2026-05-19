# Outlook MCP

A Model Context Protocol (MCP) server that exposes Microsoft Outlook **mail** and **calendar** to AI agents via the Microsoft Graph SDK. Acts as the signed-in user (delegated permissions, MSAL device code flow).

## What it does

29 workflow-oriented tools covering the common mail and calendar operations:

| Group | Tools |
|---|---|
| Util | `whoami` |
| Mail — read | `list_messages`, `search_messages`, `get_message`, `list_attachments`, `download_attachment` |
| Mail — write | `send_message`, `create_draft`, `reply_message`, `reply_all_message`, `forward_message`, `update_message`, `delete_message` |
| Mail — folders | `list_folders`, `move_message` |
| Mail — actions | `archive_message`, `mark_read`, `mark_unread`, `flag_message`, `unflag_message` |
| Calendar | `list_calendars`, `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `cancel_event`, `respond_to_event`, `find_meeting_times` |

Every tool that touches a mailbox or calendar accepts an optional `mailbox` argument (email or user ID) to target shared mailboxes/calendars. Omit it to use the signed-in user's own mailbox.

Every tool that returns objects accepts `include_raw=true` to also include the full Graph payload.

List/search tools support pagination via `limit` (1-100, default 25) and `page_token`.

## Prerequisites

- Python ≥ 3.11
- `uv`: https://docs.astral.sh/uv/
- An Entra (Azure AD) app registration with the right permissions (see "Entra setup" below)

## Quickstart

```bash
# 1. Install dependencies
uv sync

# 2. Configure environment
cp .env.example .env
# Fill in OUTLOOK_MCP_CLIENT_ID and OUTLOOK_MCP_TENANT_ID

# 3. One-time sign-in (device code flow)
uv run outlook-mcp-login
# Follow the prompt: visit the URL, enter the code, complete sign-in.
# A token cache is written to ~/.outlook-mcp/token_cache.bin (mode 0600).

# 4. Wire the MCP into your host
# - Claude Code:
claude mcp add outlook -- uv --directory "$(pwd)" run outlook-mcp

# - Anything else: configure the host to launch `uv run outlook-mcp` (stdio).
```

## Entra setup

The app registration ("Outlook MCP" in the Sentasity tenant) requires:

- **Account type:** single tenant
- **Redirect URI (public client):** `https://login.microsoftonline.com/common/oauth2/nativeclient`
- **Delegated permissions** (Microsoft Graph):
  - `Mail.ReadWrite`
  - `Mail.ReadWrite.Shared`
  - `Mail.Send`
  - `Calendars.ReadWrite`
  - `Calendars.ReadWrite.Shared`
  - `User.Read`
- Admin consent for the Shared permissions if your tenant requires it.

The CLI uses public-client device code flow — **no client secret** is needed or stored.

## Environment variables

| Var | Required | Default | Purpose |
|---|---|---|---|
| `OUTLOOK_MCP_CLIENT_ID` | yes | — | Entra (Azure AD) app client ID |
| `OUTLOOK_MCP_TENANT_ID` | yes | — | Tenant ID (single-tenant authority) |
| `OUTLOOK_MCP_TOKEN_CACHE_PATH` | no | `~/.outlook-mcp/token_cache.bin` | Override token cache file location |

Process env wins; `.env` at the repo root is loaded as a dev fallback.

## Security

- The token cache contains your **refresh token**, which can mint access tokens for your Outlook data. Treat it like a credential.
- Default location: `~/.outlook-mcp/token_cache.bin`, mode `0600`, parent dir mode `0700`.
- To **revoke** access: sign in to https://account.microsoft.com or your org's identity portal, revoke the "Outlook MCP" app, then `rm ~/.outlook-mcp/token_cache.bin`.
- To **switch accounts**: `rm ~/.outlook-mcp/token_cache.bin` and re-run `outlook-mcp-login`.

## Future work

Not implemented in v1; reasonable v2 additions:

- **Rules and inbox automation** (`messageRules`). Useful for codifying inbox triage, less common in agent workflows.
- **Graph `$batch` requests.** Performance optimization that bundles multiple Graph calls into one HTTP round-trip; would speed up multi-step workflows but adds complexity. Defer until profiling proves the win.

Other known gaps (intentionally out of scope for v1): chunked attachment upload (>3 MB), category master-list management, mail signatures, contacts/To-Do/OneNote/Teams, multi-account switching, change-notification subscriptions.

## Development

```bash
# Run unit tests
uv run pytest

# Run unit + live integration smoke (requires a valid token cache)
OUTLOOK_MCP_INTEGRATION=1 uv run pytest

# Type check
uv run pyright src tests

# Lint
uv run ruff check .
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| `NotAuthenticatedError: Not authenticated. Run \`outlook-mcp-login\`...` | Run `uv run outlook-mcp-login`. |
| `ConfigError: Missing required env var: OUTLOOK_MCP_CLIENT_ID` | Set the var in `.env` or in your MCP host's env config. |
| `Graph API 403: ErrorAccessDenied — ...` | Permission mismatch on the Entra app. Verify the delegated permissions list above and re-consent. |
| Server boots but tools 404 in the host | Confirm the host is launching `uv run outlook-mcp` with the right working directory. |

---

Spec: `docs/specs/2026-05-19-outlook-mcp-design.md` (gitignored — local working doc).
