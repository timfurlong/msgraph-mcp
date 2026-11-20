package tools

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

type messageActionIn struct {
	MessageID  string  `json:"message_id"`
	Mailbox    *string `json:"mailbox"`
	IncludeRaw bool    `json:"include_raw"`
}

func registerMailActions(s *mcp.Server, g *graph.Client) {
	add(s, "archive_message", func(ctx context.Context, in messageActionIn) (any, error) {
		return moveMessage(ctx, g, moveMessageIn{MessageID: in.MessageID, Destination: "archive", Mailbox: in.Mailbox, IncludeRaw: in.IncludeRaw})
	})
	update := func(name string, set func(*updateMessageIn)) {
		add(s, name, func(ctx context.Context, in messageActionIn) (any, error) {
			u := updateMessageIn{MessageID: in.MessageID, Mailbox: in.Mailbox, IncludeRaw: in.IncludeRaw}
			set(&u)
			return updateMessage(ctx, g, u)
		})
	}
	t, f := true, false
	flagged, notFlagged := "flagged", "notFlagged"
	update("mark_read", func(u *updateMessageIn) { u.IsRead = &t })
	update("mark_unread", func(u *updateMessageIn) { u.IsRead = &f })
	update("flag_message", func(u *updateMessageIn) { u.Flag = &flagged })
	update("unflag_message", func(u *updateMessageIn) { u.Flag = &notFlagged })
}
