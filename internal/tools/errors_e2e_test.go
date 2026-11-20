package tools

import (
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func TestToolErrorsEndToEnd(t *testing.T) {
	cases := map[string]M{
		"list_messages":   nil,
		"search_messages": {"query": "x"},
		"get_message":     {"message_id": "m1"},
		"list_folders":    nil,
		"list_rules":      nil,
		"list_events":     {"calendar_id": "c", "start_datetime": "a", "end_datetime": "b"},
		"list_chats":      nil,
		"whoami":          nil,
		"batch_mark_read": {"message_ids": []any{"m1"}},
	}
	for tool, args := range cases {
		cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
			return graphtest.Error(403, "ErrorAccessDenied", "Access is denied.")
		})
		if msg := callErr(t, cs, tool, args); msg != "Graph API 403: ErrorAccessDenied — Access is denied." {
			t.Fatalf("%s: %q", tool, msg)
		}
	}
}
