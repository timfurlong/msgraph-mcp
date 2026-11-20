package tools

import (
	"fmt"
	"strings"
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

// batchEcho answers every sub-request 200, except ids listed in fail (404).
func batchEcho(fail map[string]bool) graphtest.Handler {
	return func(r graphtest.Request) graphtest.Response {
		var out []any
		for _, sub := range r.Body["requests"].([]any) {
			m := sub.(map[string]any)
			status := 200
			body := any(M{"id": m["id"]})
			for id := range fail {
				if url, _ := m["url"].(string); len(url) > 0 && containsID(url, id) {
					status = 404
					body = M{"error": M{"code": "ErrorItemNotFound", "message": "gone"}}
				}
			}
			out = append(out, M{"id": m["id"], "status": status, "body": body})
		}
		return jsonOK(M{"responses": out})
	}
}

func containsID(url, id string) bool {
	return strings.HasSuffix(url, "/"+id) || strings.Contains(url, "/"+id+"/")
}

func TestBatchTools(t *testing.T) {
	cases := []struct {
		tool   string
		extra  M
		method string
		suffix string
		body   string
	}{
		{"batch_archive_messages", nil, "POST", "/move", `{"destinationId":"archive"}`},
		{"batch_move_messages", M{"destination": "f1"}, "POST", "/move", `{"destinationId":"f1"}`},
		{"batch_mark_read", nil, "PATCH", "", `{"isRead":true}`},
		{"batch_mark_unread", nil, "PATCH", "", `{"isRead":false}`},
		{"batch_flag_messages", nil, "PATCH", "", `{"flag":{"flagStatus":"flagged"}}`},
		{"batch_unflag_messages", nil, "PATCH", "", `{"flag":{"flagStatus":"notFlagged"}}`},
	}
	for _, c := range cases {
		cs, f := newSession(t, batchEcho(map[string]bool{"m2": true}))
		args := M{"message_ids": []any{"m1", "m2"}, "mailbox": "a+b@example.com"}
		for k, v := range c.extra {
			args[k] = v
		}
		got := callJSON(t, cs, c.tool, args)
		eq(t, c.tool+" result", got, M{
			"results": []any{M{"id": "m1", "ok": true}, M{"id": "m2", "ok": false, "error": "404 ErrorItemNotFound: gone"}},
			"summary": M{"total": 2, "succeeded": 1, "failed": 1},
		})
		sub := f.Last().Body["requests"].([]any)[0].(map[string]any)
		if sub["method"] != c.method || sub["url"] != "/users/a+b@example.com/messages/m1"+c.suffix {
			t.Fatalf("%s: %v", c.tool, sub)
		}
		eq(t, c.tool+" body", sub["body"], mustJSON(c.body))
	}
}

func mustJSON(s string) any {
	var v any
	if err := jsonUnmarshal(s, &v); err != nil {
		panic(err)
	}
	return v
}

func TestBatchToolsValidation(t *testing.T) {
	cs, _ := newSession(t, nil)
	if msg := callErr(t, cs, "batch_mark_read", M{"message_ids": []any{}}); msg != "`message_ids` must be a non-empty list" {
		t.Fatalf("msg %q", msg)
	}
	ids := []any{}
	for i := 0; i < 1001; i++ {
		ids = append(ids, fmt.Sprint(i))
	}
	if msg := callErr(t, cs, "batch_mark_read", M{"message_ids": ids}); msg != "`message_ids` exceeds maximum of 1000 per call" {
		t.Fatalf("msg %q", msg)
	}
	if msg := callErr(t, cs, "batch_move_messages", M{"message_ids": []any{"m1"}, "destination": ""}); msg != "`destination` is required" {
		t.Fatalf("msg %q", msg)
	}
}

func TestBatchToolsMeRoot(t *testing.T) {
	cs, f := newSession(t, batchEcho(nil))
	callJSON(t, cs, "batch_mark_read", M{"message_ids": []any{"m1"}})
	if url := f.Last().Body["requests"].([]any)[0].(map[string]any)["url"]; url != "/me/messages/m1" {
		t.Fatalf("url %v", url)
	}
}

var ruleJSON = M{"id": "r1", "displayName": "Notify", "sequence": 1, "isEnabled": true,
	"conditions": M{"senderContains": []any{"example.com"}}, "actions": M{"moveToFolder": "F", "stopProcessingRules": true}}

func TestListAndGetRules(t *testing.T) {
	cs, f := newSession(t, func(r graphtest.Request) graphtest.Response {
		if r.Path == "/me/mailFolders/inbox/messageRules" {
			return jsonOK(M{"value": []any{ruleJSON}})
		}
		return jsonOK(ruleJSON)
	})
	got := callJSON(t, cs, "list_rules", nil)
	if got["items"].([]any)[0].(M)["display_name"] != "Notify" || got["next_page_token"] != nil {
		t.Fatalf("got %v", got)
	}
	got = callJSON(t, cs, "get_rule", M{"rule_id": "r1"})
	if f.Last().Path != "/me/mailFolders/inbox/messageRules/r1" || got["id"] != "r1" {
		t.Fatalf("%s %v", f.Last().Path, got)
	}
	if msg := callErr(t, cs, "get_rule", M{"rule_id": ""}); msg != "`rule_id` is required" {
		t.Fatalf("msg %q", msg)
	}
}

func TestCreateRule(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 201, JSON: ruleJSON} })
	got := callJSON(t, cs, "create_rule", M{"display_name": "Notify", "sender_contains": []any{"example.com"}, "from_addresses": []any{"a@example.com"}, "move_to_folder": "F", "stop_processing_rules": true})
	if got["id"] != "r1" {
		t.Fatalf("got %v", got)
	}
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/mailFolders/inbox/messageRules" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "rule body", r.Body, M{
		"displayName": "Notify", "isEnabled": true, "sequence": 1,
		"conditions": M{"senderContains": []any{"example.com"}, "fromAddresses": recips("a@example.com")},
		"actions":    M{"moveToFolder": "F", "stopProcessingRules": true},
	})

	callJSON(t, cs, "create_rule", M{"display_name": "X", "has_attachments": false, "delete": true, "sequence": 5, "is_enabled": false})
	eq(t, "rule body 2", f.Last().Body, M{"displayName": "X", "isEnabled": false, "sequence": 5, "conditions": M{"hasAttachments": false}, "actions": M{"delete": true}})
}

func TestCreateRuleValidation(t *testing.T) {
	cs, _ := newSession(t, nil)
	for want, args := range map[string]M{
		"`display_name` is required": {"display_name": " ", "sender_contains": []any{"x"}, "delete": true},
		"At least one condition is required (sender_contains, subject_contains, body_contains, body_or_subject_contains, from_addresses, or has_attachments)": {"display_name": "n", "sender_contains": []any{}, "delete": true},
		"At least one action is required (move_to_folder, mark_as_read, delete, or stop_processing_rules)":                                                    {"display_name": "n", "sender_contains": []any{"x"}},
		"Invalid email address: 'nobody'": {"display_name": "n", "from_addresses": []any{"nobody"}, "delete": true},
	} {
		if msg := callErr(t, cs, "create_rule", args); msg != want {
			t.Fatalf("msg %q, want %q", msg, want)
		}
	}
}

func TestUpdateRule(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(ruleJSON) })
	callJSON(t, cs, "update_rule", M{"rule_id": "r1", "is_enabled": false})
	r := f.Last()
	if r.Method != "PATCH" || r.Path != "/me/mailFolders/inbox/messageRules/r1" || string(r.RawBody) != `{"isEnabled":false}` {
		t.Fatalf("%s %s %s", r.Method, r.Path, r.RawBody)
	}
	callJSON(t, cs, "update_rule", M{"rule_id": "r1", "subject_contains": []any{"x"}, "mark_as_read": true, "sequence": 3, "display_name": "D"})
	eq(t, "patch", f.Last().Body, M{"displayName": "D", "sequence": 3, "conditions": M{"subjectContains": []any{"x"}}, "actions": M{"markAsRead": true}})
	for want, args := range map[string]M{
		"`rule_id` is required": {"rule_id": "", "is_enabled": true},
		"At least one field must be updated (display_name, is_enabled, sequence, a condition, or an action)": {"rule_id": "r1"},
		"`display_name` cannot be empty": {"rule_id": "r1", "display_name": " "},
	} {
		if msg := callErr(t, cs, "update_rule", args); msg != want {
			t.Fatalf("msg %q, want %q", msg, want)
		}
	}
}

func TestDeleteRule(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 204} })
	got := callJSON(t, cs, "delete_rule", M{"rule_id": "r1"})
	eq(t, "result", got, M{"deleted": true, "id": "r1"})
	if f.Last().Method != "DELETE" || f.Last().Path != "/me/mailFolders/inbox/messageRules/r1" {
		t.Fatalf("%s %s", f.Last().Method, f.Last().Path)
	}
}
