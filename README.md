# MSGraph MCP

<!-- mcp-name: io.github.timfurlong/msgraph-mcp -->

[![PyPI](https://img.shields.io/pypi/v/msgraph-mcp-server)](https://pypi.org/project/msgraph-mcp-server/)
[![Python versions](https://img.shields.io/pypi/pyversions/msgraph-mcp-server)](https://pypi.org/project/msgraph-mcp-server/)
[![CI](https://github.com/timfurlong/msgraph-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/timfurlong/msgraph-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/timfurlong/msgraph-mcp/blob/main/LICENSE)

A Model Context Protocol (MCP) server for **Microsoft Graph**. It exposes Microsoft Outlook **mail** and **calendar**, plus read-only Microsoft **Teams** message history, to AI agents via the Microsoft Graph SDK. It acts as the signed-in user, using delegated permissions and MSAL device code flow, so it can only reach what that user can reach.

## Install

You need Python ≥ 3.11 and an Entra (Azure AD) app registration. The app registration takes about ten minutes and may need an administrator, so do that first: see the Entra setup section below.

```bash
uv tool install msgraph-mcp-server   # or: pip install msgraph-mcp-server
```

The package is `msgraph-mcp-server`. It installs three commands: `msgraph-mcp` and its alias `msgraph-mcp-server`, which both start the stdio server, plus `msgraph-mcp-login` for the one-time sign-in.

## Setup

**1. Set your Entra app credentials.**

```bash
export MSGRAPH_MCP_CLIENT_ID=<your app's client ID>
export MSGRAPH_MCP_TENANT_ID=<your tenant ID>   # or: common / organizations / consumers
```

**2. Sign in once.** This prints a URL and a code; visit the URL and enter the code. No client secret is involved or stored.

```bash
msgraph-mcp-login
# running via uvx instead of installing: uvx --from msgraph-mcp-server msgraph-mcp-login
```

A token cache is written to `~/.msgraph-mcp/token_cache.bin`.

**3. Wire the server into your MCP host.**

Claude Code:

```bash
claude mcp add msgraph \
  --env MSGRAPH_MCP_CLIENT_ID=$MSGRAPH_MCP_CLIENT_ID \
  --env MSGRAPH_MCP_TENANT_ID=$MSGRAPH_MCP_TENANT_ID \
  -- msgraph-mcp
```

Any other host, over stdio:

```json
{
  "mcpServers": {
    "msgraph": {
      "command": "msgraph-mcp",
      "args": [],
      "env": {
        "MSGRAPH_MCP_CLIENT_ID": "<your app's client ID>",
        "MSGRAPH_MCP_TENANT_ID": "<your tenant ID>"
      }
    }
  }
}
```

To run without installing, use `"command": "uvx"` with `"args": ["msgraph-mcp-server"]`. Note that `uvx` does not put `msgraph-mcp-login` on your `PATH`, so sign in with the `uvx --from` form shown in step 2.

`MSGRAPH_MCP_TOKEN_CACHE_PATH` optionally overrides the cache location. All three variables can also come from a `.env` file in the working directory; the process environment wins over it.

## Entra setup

The app registration requires:

- **Account type:** single tenant (multi-tenant works too; set `MSGRAPH_MCP_TENANT_ID` to `common`, `organizations`, or `consumers`)
- **Redirect URI (public client):** `https://login.microsoftonline.com/common/oauth2/nativeclient`
- **Allow public client flows:** Yes, under Authentication → Advanced settings. Device code flow fails without it.
- **Delegated permissions** (Microsoft Graph):
  - Mail: `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `Mail.Send`
  - Rules: `MailboxSettings.ReadWrite` (Graph requires this for the `messageRules` endpoints)
  - Calendar: `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`
  - Identity: `User.Read`
  - Teams: `Chat.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`

`ChannelMessage.Read.All` always needs tenant admin consent, and the `*.Shared` permissions may need it depending on your tenant.

> **The scope list is all-or-nothing.** Sign-in requests every scope at once, so without admin consent for `ChannelMessage.Read.All` the login fails outright and mail and calendar are unavailable too. For the same reason, adding scopes later means re-running `msgraph-mcp-login`; until you do, *every* tool fails with `NotAuthenticatedError`, not just the ones needing the new scope.

## Tools

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

Behavior shared across tools:

- **Trimmed responses.** Results are reshaped for agents, and message and event bodies are replaced by a short `snippet`. Pass `include_body=true` to `get_message`, `list_chat_messages`, `list_channel_messages`, or `list_message_replies` for the full body, which for Teams is also what surfaces Adaptive Card content. Pass `include_raw=true` to any tool returning a Graph object to get the full payload alongside the trimmed one; the `batch_*` tools return per-message status only and do not accept it.
- **Other mailboxes.** Every mail and calendar tool takes an optional `mailbox` (email or user ID) to target a shared or delegated mailbox. Omit it for your own. Teams tools are read-only, cover only your own chats, and take no `mailbox`.
- **Pagination.** List and search tools take `limit` (1-100, default 25) and `page_token`, except `list_channel_messages` and `list_message_replies`, which Graph caps at 50. `list_folders`, `list_rules`, `list_attachments`, and `list_calendars` return the whole collection and take neither. On `list_joined_teams` and `list_channels`, `limit` is applied after fetching because Graph rejects `$top` there, so it saves tokens rather than round-trips.
- **Batch actions.** The `batch_*` tools apply one action to up to 1000 messages via Graph's `$batch` endpoint, returning a per-message result plus a `{total, succeeded, failed}` summary.
- **Downloads.** `download_attachment` and `download_hosted_content` return images as native MCP image blocks the agent can view. Pass `save_path` (a file, or an existing directory) to write bytes to disk and get back a path instead.
- **Finding mail.** `list_messages` defaults to the inbox; pass `folder_id` for another folder, `unread_only=true`, or a raw OData `filter` for predicates KQL cannot express. `search_messages` passes your query to Graph's `$search` as KQL.

Outgoing attachments are capped at 3 MB total per message; chunked upload is not supported. SharePoint-backed Teams file attachments cannot be downloaded, only inline hosted content.

## Security

The token cache holds your **refresh token**, which can mint access tokens for your mail, calendar, and Teams data. Treat the file as a credential. It lives at `~/.msgraph-mcp/token_cache.bin` with mode `0600` inside a `0700` directory.

To revoke access, or to switch accounts, delete the cache and sign in again:

```bash
rm -f ~/.msgraph-mcp/token_cache.bin ~/.outlook-mcp/token_cache.bin
```

The second path matters if you ever ran this server under its former name `outlook-mcp`: that cache is still used as a fallback when the current one is absent, so deleting only the first file leaves a working refresh token on disk. To revoke fully, also remove the app at https://account.microsoft.com or in your organization's identity portal.

## Troubleshooting

- **`NotAuthenticatedError: Not authenticated`** — run `msgraph-mcp-login`. If it recurs immediately, a requested scope has not been consented yet; see the note under Entra setup.
- **`ConfigError: Missing required env var`** — set it in your shell, in a `.env` in the working directory, or in your MCP host's env config.
- **`Graph API 403: ErrorAccessDenied`** — the Entra app is missing a delegated permission, or it needs admin consent. Check the list above, then re-consent and sign in again.
- **`Graph API 400: BadRequest — Syntax error`** from `search_messages` — the query goes to Graph's `$search` as KQL. Quote literal phrases (`"weekly report"`) or use fielded forms (`from:alice`). For predicates KQL cannot express, use `list_messages` with `filter=`.
- **The host starts the server but lists no tools** — confirm it launches `msgraph-mcp` (or `uvx`) over stdio and can find that binary on its `PATH`.

## Development

See [CONTRIBUTING.md](https://github.com/timfurlong/msgraph-mcp/blob/main/CONTRIBUTING.md).
