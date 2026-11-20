package tools

import (
	"encoding/base64"
	"reflect"
	"strings"
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func accepted(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 202} }

func recips(addrs ...string) []any {
	out := []any{}
	for _, a := range addrs {
		out = append(out, M{"emailAddress": M{"address": a}})
	}
	return out
}

func eq(t *testing.T, what string, got, want any) {
	t.Helper()
	g, _ := jsonText(got)
	w, _ := jsonText(want)
	if g != w {
		t.Fatalf("%s:\n got %s\nwant %s", what, g, w)
	}
}

func TestSendMessageMinimal(t *testing.T) {
	cs, f := newSession(t, accepted)
	got := callJSON(t, cs, "send_message", M{"to": []any{"bob@example.com"}, "subject": "Hi", "body": "Hello"})
	eq(t, "result", got, M{"status": "sent"})
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/sendMail" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "body", r.Body, M{
		"message": M{
			"subject":       "Hi",
			"body":          M{"contentType": "text", "content": "Hello"},
			"toRecipients":  recips("bob@example.com"),
			"ccRecipients":  []any{},
			"bccRecipients": []any{},
		},
		"saveToSentItems": true,
	})
}

func TestSendMessageHTMLCcBccAttachmentsMailbox(t *testing.T) {
	cs, f := newSession(t, accepted)
	callJSON(t, cs, "send_message", M{
		"to": []any{"a@example.com"}, "cc": []any{"c@example.com"}, "bcc": []any{"b@example.com"},
		"subject": "S", "body": "<p>x</p>", "body_type": "html", "save_to_sent_items": false, "mailbox": "shared@example.com",
		"attachments": []any{M{"name": "a.txt", "content_type": "text/plain", "content_b64": "YWJj"}, M{"content_b64": "eHl6"}},
	})
	r := f.Last()
	if r.RawPath != "/users/shared%40example.com/sendMail" || r.Body["saveToSentItems"] != false {
		t.Fatalf("%s %v", r.RawPath, r.Body["saveToSentItems"])
	}
	msg := r.Body["message"].(map[string]any)
	eq(t, "body", msg["body"], M{"contentType": "html", "content": "<p>x</p>"})
	eq(t, "cc", msg["ccRecipients"], recips("c@example.com"))
	eq(t, "attachments", msg["attachments"], []any{
		M{"@odata.type": "#microsoft.graph.fileAttachment", "name": "a.txt", "contentType": "text/plain", "contentBytes": "YWJj"},
		M{"@odata.type": "#microsoft.graph.fileAttachment", "name": "attachment", "contentType": "application/octet-stream", "contentBytes": "eHl6"},
	})
}

func TestSendMessageRequiresTo(t *testing.T) {
	cs, f := newSession(t, accepted)
	if msg := callErr(t, cs, "send_message", M{"to": []any{}, "subject": "s", "body": "b"}); msg != "`to` must be a non-empty list of email addresses" {
		t.Fatalf("msg %q", msg)
	}
	if len(f.Requests) != 0 {
		t.Fatal("called Graph")
	}
}

func TestSendMessageRejectsOversizeAttachment(t *testing.T) {
	cs, f := newSession(t, accepted)
	big := base64.StdEncoding.EncodeToString(make([]byte, 3*1024*1024+1))
	msg := callErr(t, cs, "send_message", M{"to": []any{"a@example.com"}, "subject": "s", "body": "b", "attachments": []any{M{"name": "big", "content_b64": big}}})
	if msg != "Attachment payload exceeds 3 MB raw limit (got 3145729 bytes). Use a chunked upload (not supported in v1) for larger files." {
		t.Fatalf("msg %q", msg)
	}
	if len(f.Requests) != 0 {
		t.Fatal("called Graph")
	}
}

func TestSendMessageRejectsBadBase64(t *testing.T) {
	cs, _ := newSession(t, accepted)
	msg := callErr(t, cs, "send_message", M{"to": []any{"a@example.com"}, "subject": "s", "body": "b", "attachments": []any{M{"content_b64": "not base64!"}}})
	if !strings.HasPrefix(msg, "Invalid attachment content_b64: ") {
		t.Fatalf("msg %q", msg)
	}
	msg = callErr(t, cs, "send_message", M{"to": []any{"a@example.com"}, "subject": "s", "body": "b", "attachments": []any{M{"name": "x"}}})
	if !strings.HasPrefix(msg, "Invalid attachment content_b64: ") {
		t.Fatalf("msg %q", msg)
	}
}

func TestCreateDraftReturnsTrimmed(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Status: 201, JSON: M{"id": "d1", "subject": "Draft", "isDraft": true}}
	})
	got := callJSON(t, cs, "create_draft", M{"to": []any{"bob@example.com"}, "subject": "Draft", "body": "x"})
	if got["id"] != "d1" || got["is_draft"] != true {
		t.Fatalf("got %v", got)
	}
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/messages" || r.Body["subject"] != "Draft" {
		t.Fatalf("%s %s %v", r.Method, r.Path, r.Body)
	}
	if !reflect.DeepEqual(r.Body["toRecipients"], recips("bob@example.com")) {
		t.Fatalf("to %v", r.Body["toRecipients"])
	}
}

func TestReplyMessage(t *testing.T) {
	cs, f := newSession(t, accepted)
	got := callJSON(t, cs, "reply_message", M{"message_id": "m1", "comment": "Thanks", "extra_to": []any{"x@example.com"}})
	eq(t, "result", got, M{"status": "sent"})
	r := f.Last()
	if r.Path != "/me/messages/m1/reply" {
		t.Fatalf("path %s", r.Path)
	}
	eq(t, "body", r.Body, M{"comment": "Thanks", "message": M{"toRecipients": recips("x@example.com")}})

	callJSON(t, cs, "reply_message", M{"message_id": "m1"})
	if string(f.Last().RawBody) != "{}" {
		t.Fatalf("bare reply body %s", f.Last().RawBody)
	}
}

func TestReplyAllMessage(t *testing.T) {
	cs, f := newSession(t, accepted)
	got := callJSON(t, cs, "reply_all_message", M{"message_id": "m1", "comment": "All"})
	eq(t, "result", got, M{"status": "sent"})
	if r := f.Last(); r.Path != "/me/messages/m1/replyAll" || r.Body["comment"] != "All" {
		t.Fatalf("%s %v", r.Path, r.Body)
	}
}

func TestForwardMessage(t *testing.T) {
	cs, f := newSession(t, accepted)
	got := callJSON(t, cs, "forward_message", M{"message_id": "m1", "to": []any{"f@example.com"}})
	eq(t, "result", got, M{"status": "sent"})
	r := f.Last()
	if r.Path != "/me/messages/m1/forward" {
		t.Fatalf("path %s", r.Path)
	}
	eq(t, "body", r.Body, M{"toRecipients": recips("f@example.com")})
	if msg := callErr(t, cs, "forward_message", M{"message_id": "m1", "to": []any{}}); msg != "`to` must be a non-empty list of email addresses" {
		t.Fatalf("msg %q", msg)
	}
}

func TestUpdateMessage(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"id": "m1", "isRead": true})
	})
	got := callJSON(t, cs, "update_message", M{"message_id": "m1", "is_read": true, "flag": "flagged", "importance": "high", "categories": []any{"Blue"}, "parent_folder_id": "f1"})
	if got["is_read"] != true {
		t.Fatalf("got %v", got)
	}
	r := f.Last()
	if r.Method != "PATCH" || r.Path != "/me/messages/m1" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
	eq(t, "patch", r.Body, M{"isRead": true, "flag": M{"flagStatus": "flagged"}, "importance": "high", "categories": []any{"Blue"}, "parentFolderId": "f1"})

	callJSON(t, cs, "update_message", M{"message_id": "m1", "is_read": false})
	if string(f.Last().RawBody) != `{"isRead":false}` {
		t.Fatalf("patch %s", f.Last().RawBody)
	}
}

func TestDeleteMessage(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 204} })
	got := callJSON(t, cs, "delete_message", M{"message_id": "m1"})
	eq(t, "result", got, M{"status": "deleted"})
	if r := f.Last(); r.Method != "DELETE" || r.Path != "/me/messages/m1" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
}
