package toolspec

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestSnapshotHas50Tools(t *testing.T) {
	all := All()
	if len(all) != 50 {
		t.Fatalf("got %d tools, want 50", len(all))
	}
	if all[0].Name != "whoami" {
		t.Fatalf("first tool = %q, want whoami", all[0].Name)
	}
	seen := map[string]bool{}
	for _, tool := range all {
		if seen[tool.Name] {
			t.Fatalf("duplicate tool %q", tool.Name)
		}
		seen[tool.Name] = true
		if tool.Description == "" {
			t.Fatalf("%s has no description", tool.Name)
		}
	}
}

func TestEveryInputSchemaIsObject(t *testing.T) {
	for _, tool := range All() {
		var s struct {
			Type string `json:"type"`
		}
		if err := json.Unmarshal(tool.InputSchema, &s); err != nil {
			t.Fatalf("%s: %v", tool.Name, err)
		}
		if s.Type != "object" {
			t.Fatalf("%s: schema type %q, want object", tool.Name, s.Type)
		}
	}
}

func TestListChatMessagesLimitDocumented50(t *testing.T) {
	tool, ok := Get("list_chat_messages")
	if !ok {
		t.Fatal("list_chat_messages missing")
	}
	if !strings.Contains(tool.Description, "limit: 1-50.") {
		t.Fatalf("description does not document the 50 cap:\n%s", tool.Description)
	}
	if _, ok := Get("no_such_tool"); ok {
		t.Fatal("Get found a tool that does not exist")
	}
}
