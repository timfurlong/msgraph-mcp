# MSGraph MCP

<!-- mcp-name: io.github.timfurlong/msgraph-mcp -->

[![PyPI](https://img.shields.io/pypi/v/msgraph-mcp-server)](https://pypi.org/project/msgraph-mcp-server/)
[![Python versions](https://img.shields.io/pypi/pyversions/msgraph-mcp-server)](https://pypi.org/project/msgraph-mcp-server/)
[![CI](https://github.com/timfurlong/msgraph-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/timfurlong/msgraph-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A Model Context Protocol (MCP) server for **Microsoft Graph**. It exposes Microsoft Outlook **mail** and **calendar**, plus read-only Microsoft **Teams** message history, to AI agents via the Microsoft Graph SDK. Acts as the signed-in user (delegated permissions, MSAL device code flow).

## What it does

Workflow-oriented tools covering common mail, calendar, and read-only Teams operations:

| Group          | Tools                                                                                                                                                  |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Util           | `whoami`                                                                                                                                               |
| Mail — read    | `list_messages`, `search_messages`, `get_message`, `list_attachments`, `download_attachment`                                                           |
| Mail — write   | `send_message`, `create_draft`, `reply_message`, `reply_all_message`, `forward_message`, `update_message`, `delete_message`                            |
| Mail — folders | `list_folders`, `create_folder`, `update_folder`, `delete_folder`, `move_message`                                                                      |
| Mail — actions | `archive_message`, `mark_read`, `mark_unread`, `flag_message`, `unflag_message`                                                                        |
| Mail — batch   | `batch_archive_messages`, `batch_move_messages`, `batch_mark_read`, `batch_mark_unread`, `batch_flag_messages`, `batch_unflag_messages`                |
| Mail — rules   | `list_rules`, `get_rule`, `create_rule`, `update_rule`, `delete_rule`                                                                                  |
| Calendar       | `list_calendars`, `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `cancel_event`, `respond_to_event`, `find_meeting_times` |
| Teams (read)   | `list_chats`, `list_chat_messages`, `list_joined_teams`, `list_channels`, `list_channel_messages`, `list_message_replies`, `download_hosted_content`      |

Every tool that touches a mailbox or calendar accepts an optional `mailbox` argument (email or user ID) to target shared mailboxes/calendars. Omit it to use the signed-in user's own mailbox.

The `batch_*` tools apply one action to up to 1000 messages through Graph's `$batch` endpoint, and return a per-message result plus a `{total, succeeded, failed}` summary.

Every tool that returns objects accepts `include_raw=true` to also include the full Graph payload.

List/search tools support pagination via `limit` (1-100, default 25) and `page_token`.

## Install

The package is `msgraph-mcp-server`; it installs two commands, `msgraph-mcp` (the server) and `msgraph-mcp-login` (one-time sign-in).

```bash
uv tool install msgraph-mcp-server   # or: pip install msgraph-mcp-server
```

You need Python ≥ 3.11 and an Entra (Azure AD) app registration (see [Entra setup](#entra-setup)).

## Setup

**1. Set your Entra app credentials.**

```bash
export MSGRAPH_MCP_CLIENT_ID=<your app's client ID>
export MSGRAPH_MCP_TENANT_ID=<your tenant ID>
```

**2. Sign in once** (device code flow — visit the URL it prints, enter the code):

```bash
msgraph-mcp-login
```

A token cache is written to `~/.msgraph-mcp/token_cache.bin` (mode `0600`).

**3. Wire the server into your MCP host.**

Claude Code:

```bash
claude mcp add msgraph \
  --env MSGRAPH_MCP_CLIENT_ID=$MSGRAPH_MCP_CLIENT_ID \
  --env MSGRAPH_MCP_TENANT_ID=$MSGRAPH_MCP_TENANT_ID \
  -- uvx msgraph-mcp-server
```

Any other host — launch it over stdio:

```json
{
  "mcpServers": {
    "msgraph": {
      "command": "uvx",
      "args": ["msgraph-mcp-server"],
      "env": {
        "MSGRAPH_MCP_CLIENT_ID": "<your app's client ID>",
        "MSGRAPH_MCP_TENANT_ID": "<your tenant ID>"
      }
    }
  }
}
```

`uvx` runs the published package without installing it; if you used `uv tool install` or `pip install` above, use `msgraph-mcp` as the command instead.

## Entra setup

The app registration (e.g. "MSGraph MCP") requires:

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
  - `Chat.Read` (Teams)
  - `Team.ReadBasic.All` (Teams)
  - `Channel.ReadBasic.All` (Teams)
  - `ChannelMessage.Read.All` (Teams)
- Admin consent: required for `ChannelMessage.Read.All` (always), plus the `*.Shared` permissions if your tenant requires it.

> If you signed in before any of these scopes were added to the app (for example `MailboxSettings.ReadWrite`, or the Teams scopes), re-run `msgraph-mcp-login` so the cached token picks up the new scopes. Without them, calls needing the missing scope fail with a consent error.

The CLI uses public-client device code flow — **no client secret** is needed or stored.

## Environment variables

| Var                            | Required | Default                          | Purpose                             |
| ------------------------------ | -------- | -------------------------------- | ----------------------------------- |
| `MSGRAPH_MCP_CLIENT_ID`        | yes      | —                                | Entra (Azure AD) app client ID      |
| `MSGRAPH_MCP_TENANT_ID`        | yes      | —                                | Tenant ID (single-tenant authority) |
| `MSGRAPH_MCP_TOKEN_CACHE_PATH` | no       | `~/.msgraph-mcp/token_cache.bin` | Override token cache file location  |

Process env wins; a `.env` file in the working directory is loaded as a dev fallback. The legacy `OUTLOOK_MCP_*` name for each variable is still honored as a fallback (the `MSGRAPH_MCP_*` name wins when both are set).

## Security

- The token cache contains your **refresh token**, which can mint access tokens for your mail, calendar, and Teams data. Treat it like a credential.
- Default location: `~/.msgraph-mcp/token_cache.bin`, mode `0600`, parent dir mode `0700`.
- To **revoke** access: sign in to https://account.microsoft.com or your org's identity portal, revoke the app, then `rm ~/.msgraph-mcp/token_cache.bin`.
- To **switch accounts**: `rm ~/.msgraph-mcp/token_cache.bin` and re-run `msgraph-mcp-login`.

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

## Microsoft Teams (read-only)

Read Teams message history as the signed-in user:

- `list_chats`, `list_chat_messages`: your 1:1 and group chats.
- `list_joined_teams`, `list_channels`, `list_channel_messages`, `list_message_replies`: team channels and their threads.
- `download_hosted_content`: download an inline image referenced by a message (`hosted_content_refs`). Images come back as a native MCP image block the agent can view directly; pass `save_path` (file or existing directory) to write the bytes to disk and get back a path instead.

### Permissions and consent

These delegated scopes are required (already listed in `SCOPES`):

- `Chat.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`: user-consentable.
- `ChannelMessage.Read.All`: requires tenant administrator consent.

Setup:

1. Add the four delegated permissions to the app registration.
2. Grant tenant admin consent for `ChannelMessage.Read.All`.
3. Because the scope set changed, re-run the device-code login so the cached token carries the new scopes.

### Notes and limits

- Reading is delegated-only: you can read your own chats, not other users' chats.
- Channel message and reply pages are capped at 50 by Graph.
- SharePoint/OneDrive-backed file attachments are not downloadable here. In Teams, shared files are attachments whose `contentUrl` points into SharePoint, which is a different Graph surface (needs `Files.Read.All` / `Sites.Read.All` and the driveItem APIs). `download_hosted_content` covers inline hosted content (images), not shared files. This is deferred.

## Troubleshooting

| Symptom                                                                                                     | Fix                                                                                                                                                                                                                                                           |
| ----------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `NotAuthenticatedError: Not authenticated. Run \`msgraph-mcp-login\`...`                                    | Run `msgraph-mcp-login`.                                                                                                                                                                                                                                      |
| `ConfigError: Missing required env var: MSGRAPH_MCP_CLIENT_ID`                                              | Set the var in your shell, your `.env`, or your MCP host's env config.                                                                                                                                                                                        |
| `Graph API 403: ErrorAccessDenied — ...`                                                                    | Permission mismatch on the Entra app. Verify the delegated permissions list above and re-consent.                                                                                                                                                             |
| `Graph API 400: BadRequest — Syntax error: character ... is not valid at position N` from `search_messages` | The query is passed to Graph's `$search` as-is. Wrap literal/multi-character tokens in double quotes (e.g. `"weekly report"`), or use KQL fielded forms (e.g. `from:alice subject:"report"`). Bare alphanumeric strings with embedded digits are invalid KQL. |
| Server boots but tools 404 in the host                                                                      | Confirm the host is launching the server over stdio and that it can find the `uvx` / `msgraph-mcp` binary on its `PATH`.                                                                                                                                      |

## Development

```bash
git clone https://github.com/timfurlong/msgraph-mcp
cd msgraph-mcp
uv sync
cp .env.example .env   # fill in MSGRAPH_MCP_CLIENT_ID and MSGRAPH_MCP_TENANT_ID

uv run msgraph-mcp-login   # one-time sign-in
```

Point an MCP host at the working tree with `claude mcp add msgraph -- uv --directory "$(pwd)" run msgraph-mcp`.

```bash
# Unit tests
uv run pytest

# Unit + live integration smoke (requires a valid token cache)
MSGRAPH_MCP_INTEGRATION=1 uv run pytest

# Type check
uv run pyright src tests

# Lint
uv run ruff check .

# Format (CI checks this)
uv run ruff format .
```

Releases are tag-driven: pushing a `vX.Y.Z` tag matching the `pyproject.toml` version runs the checks, publishes to PyPI via Trusted Publishing, and updates the MCP registry entry.
