// Package server assembles the MCP server: instructions plus every tool.
package server

import (
	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/tools"
)

// Instructions is sent to clients at initialize.
const Instructions = `Microsoft Outlook (mail + calendar) and Microsoft Teams (read-only) via Microsoft Graph.

This server acts as the signed-in user (delegated auth). All tools that
touch a mailbox or calendar accept an optional ` + "`mailbox`" + ` argument (email
or user ID) to target shared mailboxes/calendars; omit it to use the
signed-in user's own mailbox.

Teams tools are read-only and act as the signed-in user: list_chats,
list_chat_messages, list_joined_teams, list_channels, list_channel_messages,
list_message_replies, and download_hosted_content. They take no ` + "`mailbox`" + `
argument (delegated auth reads only your own chats; channel messages are
team-scoped).

Downloads: download_attachment and download_hosted_content return
PNG/JPEG/GIF/WebP content as a native MCP image block (viewable directly,
no base64 decoding needed). Pass ` + "`save_path`" + ` (a file path, or an existing
directory) to write the bytes to disk and get back {path, content_type,
size_bytes} instead — use that for non-image content or when you want a
file on disk. Other content without ` + "`save_path`" + ` still returns base64.

Pagination: list/search tools accept ` + "`limit`" + ` (1-100, default 25) and
` + "`page_token`" + `. Pass the returned ` + "`next_page_token`" + ` back as ` + "`page_token`" + `
to fetch the next page.

Responses are trimmed by default. Pass ` + "`include_raw=true`" + ` to any tool
that returns objects to also get the full Graph payload.`

// New returns the server with every tool registered.
func New(g *graph.Client, version string) *mcp.Server {
	s := mcp.NewServer(&mcp.Implementation{Name: "msgraph-mcp", Version: version}, &mcp.ServerOptions{Instructions: Instructions})
	tools.Register(s, g)
	return s
}
