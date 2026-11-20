// Package live runs the real binary against the signed-in mailbox. It's
// skipped unless MSGRAPH_MCP_INTEGRATION=1. It only reads, plus one draft
// that it creates and deletes; it never sends anything. It logs no message
// content.
package live

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"testing"
	"time"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func session(t *testing.T) *mcp.ClientSession {
	t.Helper()
	if os.Getenv("MSGRAPH_MCP_INTEGRATION") != "1" {
		t.Skip("set MSGRAPH_MCP_INTEGRATION=1 to run against the signed-in mailbox")
	}
	_, file, _, _ := runtime.Caller(0)
	root := filepath.Join(filepath.Dir(file), "..", "..")
	bin := filepath.Join(t.TempDir(), "msgraph-mcp")
	build := exec.Command("go", "build", "-o", bin, ".")
	build.Dir = root
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build: %v\n%s", err, out)
	}
	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	t.Cleanup(cancel)
	cmd := exec.Command(bin)
	cmd.Dir = t.TempDir() // config comes from the env or ~/.msgraph-mcp/.env
	cs, err := mcp.NewClient(&mcp.Implementation{Name: "live"}, nil).Connect(ctx, &mcp.CommandTransport{Command: cmd}, nil)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { cs.Close() })
	return cs
}

func call(t *testing.T, cs *mcp.ClientSession, name string, args map[string]any) map[string]any {
	t.Helper()
	res, err := cs.CallTool(context.Background(), &mcp.CallToolParams{Name: name, Arguments: args})
	if err != nil {
		t.Fatalf("%s: %v", name, err)
	}
	text := res.Content[0].(*mcp.TextContent).Text
	if res.IsError {
		t.Fatalf("%s returned an error: %s", name, text)
	}
	var out map[string]any
	if err := json.Unmarshal([]byte(text), &out); err != nil {
		t.Fatalf("%s: not JSON", name)
	}
	return out
}

func list(t *testing.T, out map[string]any) []any {
	t.Helper()
	items, ok := out["items"].([]any)
	if !ok {
		t.Fatal("no items list")
	}
	if _, ok := out["next_page_token"]; !ok {
		t.Fatal("no next_page_token key")
	}
	return items
}

func TestLiveReads(t *testing.T) {
	cs := session(t)

	me := call(t, cs, "whoami", map[string]any{})
	if me["id"] == nil {
		t.Fatal("whoami: no id")
	}

	first := call(t, cs, "list_messages", map[string]any{"limit": 5})
	msgs := list(t, first)
	if len(msgs) == 0 {
		t.Fatal("inbox is empty")
	}
	if tok, ok := first["next_page_token"].(string); ok {
		list(t, call(t, cs, "list_messages", map[string]any{"limit": 5, "page_token": tok}))
	}
	id := msgs[0].(map[string]any)["id"].(string)
	got := call(t, cs, "get_message", map[string]any{"message_id": id, "include_body": true})
	if _, ok := got["body"]; !ok || got["id"] != id {
		t.Fatal("get_message: no body or wrong id")
	}

	list(t, call(t, cs, "search_messages", map[string]any{"query": "invoice", "limit": 3}))

	cals := list(t, call(t, cs, "list_calendars", map[string]any{}))
	if len(cals) == 0 {
		t.Fatal("no calendars")
	}
	var calID string
	for _, c := range cals {
		if c.(map[string]any)["is_default_calendar"] == true {
			calID = c.(map[string]any)["id"].(string)
		}
	}
	if calID == "" {
		calID = cals[0].(map[string]any)["id"].(string)
	}
	now := time.Now().UTC()
	list(t, call(t, cs, "list_events", map[string]any{
		"calendar_id": calID, "start_datetime": now.Format(time.RFC3339), "end_datetime": now.Add(7 * 24 * time.Hour).Format(time.RFC3339), "limit": 10,
	}))

	chats := list(t, call(t, cs, "list_chats", map[string]any{"limit": 3}))
	if len(chats) > 0 {
		chatID := chats[0].(map[string]any)["id"].(string)
		list(t, call(t, cs, "list_chat_messages", map[string]any{"chat_id": chatID, "limit": 3}))
	}
}

func TestLiveDraftRoundTrip(t *testing.T) {
	cs := session(t)
	subject := fmt.Sprintf("msgraph-mcp live test %d", time.Now().Unix())
	draft := call(t, cs, "create_draft", map[string]any{
		"to": []any{"nobody@example.com"}, "subject": subject, "body": "Created and deleted by the msgraph-mcp live suite. Never sent.",
	})
	id, _ := draft["id"].(string)
	if id == "" || draft["is_draft"] != true {
		t.Fatal("create_draft: no id or not a draft")
	}
	deleted := false
	t.Cleanup(func() {
		if !deleted {
			_, _ = cs.CallTool(context.Background(), &mcp.CallToolParams{Name: "delete_message", Arguments: map[string]any{"message_id": id}})
		}
	})
	got := call(t, cs, "get_message", map[string]any{"message_id": id})
	if got["subject"] != subject {
		t.Fatal("draft subject mismatch")
	}
	if call(t, cs, "delete_message", map[string]any{"message_id": id})["status"] != "deleted" {
		t.Fatal("delete_message")
	}
	deleted = true
}
