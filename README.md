# Outlook MCP

A Model Context Protocol (MCP) server that exposes Microsoft Outlook **mail** and **calendar** to AI agents via the Microsoft Graph SDK. Acts as the signed-in user (delegated permissions, MSAL device code flow).

## What it does

35 workflow-oriented tools covering the common mail and calendar operations:

| Group          | Tools                                                                                                                                                  |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Util           | `whoami`                                                                                                                                               |
| Mail — read    | `list_messages`, `search_messages`, `get_message`, `list_attachments`, `download_attachment`                                                           |
| Mail — write   | `send_message`, `create_draft`, `reply_message`, `reply_all_message`, `forward_message`, `update_message`, `delete_message`                            |
| Mail — folders | `list_folders`, `create_folder`, `update_folder`, `delete_folder`, `move_message`                                                                      |
| Mail — actions | `archive_message`, `mark_read`, `mark_unread`, `flag_message`, `unflag_message`                                                                        |
| Mail — rules   | `list_rules`, `get_rule`, `create_rule`, `update_rule`, `delete_rule`                                                                                  |
| Calendar       | `list_calendars`, `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `cancel_event`, `respond_to_event`, `find_meeting_times` |

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

The app registration (e.g. "Outlook MCP") requires:

- **Account type:** single tenant
- **Redirect URI (public client):** `https://login.microsoftonline.com/common/oauth2/nativeclient`
- **Delegated permissions** (Microsoft Graph):
  - `Mail.ReadWrite`
  - `Mail.ReadWrite.Shared`
  - `Mail.Send`
  - `MailboxSettings.ReadWrite`
  - `Calendars.ReadWrite`
  - `Calendars.ReadWrite.Shared`
  - `User.Read`
- Admin consent for the Shared permissions if your tenant requires it.

> If you've already signed in before `MailboxSettings.ReadWrite` was added to the scope list, re-run `uv run outlook-mcp-login` so the cached token picks up the new scope. Without it, `create_rule` and `delete_rule` will fail with a consent error.

The CLI uses public-client device code flow — **no client secret** is needed or stored.

## Environment variables

| Var                            | Required | Default                          | Purpose                             |
| ------------------------------ | -------- | -------------------------------- | ----------------------------------- |
| `OUTLOOK_MCP_CLIENT_ID`        | yes      | —                                | Entra (Azure AD) app client ID      |
| `OUTLOOK_MCP_TENANT_ID`        | yes      | —                                | Tenant ID (single-tenant authority) |
| `OUTLOOK_MCP_TOKEN_CACHE_PATH` | no       | `~/.outlook-mcp/token_cache.bin` | Override token cache file location  |

Process env wins; `.env` at the repo root is loaded as a dev fallback.

## Security

- The token cache contains your **refresh token**, which can mint access tokens for your Outlook data. Treat it like a credential.
- Default location: `~/.outlook-mcp/token_cache.bin`, mode `0600`, parent dir mode `0700`.
- To **revoke** access: sign in to https://account.microsoft.com or your org's identity portal, revoke the "Outlook MCP" app, then `rm ~/.outlook-mcp/token_cache.bin`.
- To **switch accounts**: `rm ~/.outlook-mcp/token_cache.bin` and re-run `outlook-mcp-login`.

## Recipes

### Route a sender into a new folder

```
# 1. Make a folder for the notifications.
create_folder(display_name="Notifications")
# -> {"id": "AAMkFolderId", "display_name": "Notifications", ...}

# 2. Create an inbox rule that moves matching senders into it.
create_rule(
    display_name="Notifications",
    sender_contains=["example.com"],
    move_to_folder="AAMkFolderId",
    stop_processing_rules=True,
)
```

Conditions inside one rule are AND-ed by Outlook. Pass a list to a single condition (e.g. `sender_contains=["example.com", "monitor.io"]`) for OR within that condition. Rules only run against the inbox — Graph's `messageRules` endpoint is hardcoded there and does not support per-folder rules.

`create_rule` requires at least one condition and one action. `update_rule` patches a rule in place but **replaces** the `conditions` or `actions` block whenever you pass any condition/action arg — call `get_rule` first if you need to preserve existing values.

## Future work

Not implemented; reasonable additions:

- **Graph `$batch` requests.** Performance optimization that bundles multiple Graph calls into one HTTP round-trip; would speed up multi-step workflows but adds complexity. Defer until profiling proves the win.

Other known gaps (intentionally out of scope): chunked attachment upload (>3 MB), category master-list management, mail signatures, contacts/To-Do/OneNote/Teams, multi-account switching, change-notification subscriptions, force-delete of non-empty folders.

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

| Symptom                                                                                                     | Fix                                                                                                                                                                                                                                                           |
| ----------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `NotAuthenticatedError: Not authenticated. Run \`outlook-mcp-login\`...`                                    | Run `uv run outlook-mcp-login`.                                                                                                                                                                                                                               |
| `ConfigError: Missing required env var: OUTLOOK_MCP_CLIENT_ID`                                              | Set the var in `.env` or in your MCP host's env config.                                                                                                                                                                                                       |
| `Graph API 403: ErrorAccessDenied — ...`                                                                    | Permission mismatch on the Entra app. Verify the delegated permissions list above and re-consent.                                                                                                                                                             |
| `Graph API 400: BadRequest — Syntax error: character ... is not valid at position N` from `search_messages` | The query is passed to Graph's `$search` as-is. Wrap literal/multi-character tokens in double quotes (e.g. `"weekly report"`), or use KQL fielded forms (e.g. `from:alice subject:"report"`). Bare alphanumeric strings with embedded digits are invalid KQL. |
| Server boots but tools 404 in the host                                                                      | Confirm the host is launching `uv run outlook-mcp` with the right working directory.                                                                                                                                                                          |

---

Spec: `docs/specs/2026-05-19-outlook-mcp-design.md` (gitignored — local working doc).
