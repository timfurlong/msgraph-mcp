package tools

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

const maxBatchIDs = 1000

type batchIn struct {
	MessageIDs []string `json:"message_ids"`
	Mailbox    *string  `json:"mailbox"`
}

type batchMoveIn struct {
	MessageIDs  []string `json:"message_ids"`
	Destination string   `json:"destination"`
	Mailbox     *string  `json:"mailbox"`
}

func validateIDs(ids []string) error {
	if len(ids) == 0 {
		return graph.Invalid("`message_ids` must be a non-empty list")
	}
	if len(ids) > maxBatchIDs {
		return graph.Invalid("`message_ids` exceeds maximum of %d per call", maxBatchIDs)
	}
	return nil
}

func summarize(results []graph.BatchResult) M {
	out := []any{}
	succeeded := 0
	for _, r := range results {
		entry := M{"id": r.ID, "ok": r.OK}
		if r.Error != "" {
			entry["error"] = r.Error
		}
		if r.OK {
			succeeded++
		}
		out = append(out, entry)
	}
	return M{"results": out, "summary": M{"total": len(results), "succeeded": succeeded, "failed": len(results) - succeeded}}
}

// runBatch applies one request shape to every id. Sub-request URLs carry the
// mailbox and ids unescaped, as in Python.
func runBatch(ctx context.Context, g *graph.Client, ids []string, mailbox *string, method, suffix string, body M) (any, error) {
	if err := validateIDs(ids); err != nil {
		return nil, err
	}
	base := graph.BatchRoot(mailbox)
	reqs := make([]graph.BatchRequest, len(ids))
	for i, id := range ids {
		reqs[i] = graph.BatchRequest{ID: id, Method: method, URL: base + "/messages/" + id + suffix, Body: body}
	}
	results, err := g.ExecuteBatch(ctx, reqs)
	if err != nil {
		return nil, err
	}
	return summarize(results), nil
}

func registerMailBatch(s *mcp.Server, g *graph.Client) {
	add(s, "batch_archive_messages", func(ctx context.Context, in batchIn) (any, error) {
		return runBatch(ctx, g, in.MessageIDs, in.Mailbox, "POST", "/move", M{"destinationId": "archive"})
	})
	add(s, "batch_move_messages", func(ctx context.Context, in batchMoveIn) (any, error) {
		if err := validateIDs(in.MessageIDs); err != nil {
			return nil, err
		}
		if in.Destination == "" {
			return nil, graph.Invalid("`destination` is required")
		}
		return runBatch(ctx, g, in.MessageIDs, in.Mailbox, "POST", "/move", M{"destinationId": in.Destination})
	})
	patches := []struct {
		name string
		body M
	}{
		{"batch_mark_read", M{"isRead": true}},
		{"batch_mark_unread", M{"isRead": false}},
		{"batch_flag_messages", M{"flag": M{"flagStatus": "flagged"}}},
		{"batch_unflag_messages", M{"flag": M{"flagStatus": "notFlagged"}}},
	}
	for _, p := range patches {
		add(s, p.name, func(ctx context.Context, in batchIn) (any, error) {
			return runBatch(ctx, g, in.MessageIDs, in.Mailbox, "PATCH", "", p.body)
		})
	}
}
