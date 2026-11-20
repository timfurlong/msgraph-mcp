package tools

import (
	"encoding/base64"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)


func oneMessage(graphtest.Request) graphtest.Response {
	return jsonOK(M{"value": []any{M{"id": "m1", "subject": "Hi", "isRead": false}}})
}

func TestListMessagesDefaultFolderInbox(t *testing.T) {
	cs, f := newSession(t, oneMessage)
	got := callJSON(t, cs, "list_messages", nil)
	r := f.Last()
	if r.Path != "/me/mailFolders/inbox/messages" {
		t.Fatalf("path %s", r.Path)
	}
	if r.Query.Get("$top") != "25" || r.Query.Get("$orderby") != "receivedDateTime DESC" || r.Query.Has("$filter") {
		t.Fatalf("query %v", r.Query)
	}
	items := got["items"].([]any)
	if len(items) != 1 || items[0].(M)["id"] != "m1" || got["next_page_token"] != nil {
		t.Fatalf("got %v", got)
	}
}

func TestListMessagesFolderMailboxAndNextPage(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "m1"}}, "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/alice@example.com/mailFolders/AAA/messages?$skip=5"})
	})
	got := callJSON(t, cs, "list_messages", M{"folder_id": "AAA", "mailbox": "alice@example.com", "limit": 5})
	if f.Last().RawPath != "/users/alice%40example.com/mailFolders/AAA/messages" || f.Last().Query.Get("$top") != "5" {
		t.Fatalf("request %s %v", f.Last().RawPath, f.Last().Query)
	}
	tok, _ := got["next_page_token"].(string)
	if tok == "" {
		t.Fatal("no next_page_token")
	}
	callJSON(t, cs, "list_messages", M{"page_token": tok, "unread_only": true})
	if r := f.Last(); r.Query.Get("$skip") != "5" || r.Query.Has("$filter") {
		t.Fatalf("continuation %s %v", r.Path, r.Query)
	}
}

func TestListMessagesRejectsBadLimit(t *testing.T) {
	cs, f := newSession(t, oneMessage)
	if msg := callErr(t, cs, "list_messages", M{"limit": 200}); msg != "limit must be an integer in 1..100 (got 200)" {
		t.Fatalf("msg %q", msg)
	}
	if len(f.Requests) != 0 {
		t.Fatal("called Graph")
	}
}

func TestListMessagesFilters(t *testing.T) {
	cases := []struct {
		args M
		want string
	}{
		{M{"unread_only": true}, sortClause + " and isRead eq false"},
		{M{"filter": "hasAttachments eq true"}, sortClause + " and (hasAttachments eq true)"},
		{M{"unread_only": true, "filter": "contains(subject,'report')"}, sortClause + " and isRead eq false and (contains(subject,'report'))"},
	}
	for _, c := range cases {
		cs, f := newSession(t, oneMessage)
		callJSON(t, cs, "list_messages", c.args)
		q := f.Last().Query
		if q.Get("$filter") != c.want || q.Get("$orderby") != "receivedDateTime DESC" {
			t.Fatalf("%v: query %v", c.args, q)
		}
	}
}

func TestListMessagesInvalidPageToken(t *testing.T) {
	cs, _ := newSession(t, oneMessage)
	bad := graph.EncodePageToken("https://evil.example.com/x")
	if msg := callErr(t, cs, "list_messages", M{"page_token": *bad}); msg != "Invalid page_token: must point at graph.microsoft.com (got 'evil.example.com')" {
		t.Fatalf("msg %q", msg)
	}
}

func TestSearchMessagesQuoting(t *testing.T) {
	cases := map[string]string{
		"budget review":               `"budget review"`,
		`"status: done" weekly`:       `"\"status: done\" weekly"`,
		"2fa":                         `"2fa"`,
		`from:alice subject:"report"`: `"from:alice subject:\"report\""`,
		"invoice":                     `"invoice"`,
		`C:\temp`:                     `"C:\\temp"`,
	}
	for query, sent := range cases {
		cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
			return jsonOK(M{"value": []any{M{"id": "s1"}}})
		})
		got := callJSON(t, cs, "search_messages", M{"query": query})
		r := f.Last()
		if r.Path != "/me/messages" || r.Query.Get("$search") != sent || r.Query.Get("$top") != "25" || len(r.Query) != 2 {
			t.Fatalf("%q: %s %v", query, r.Path, r.Query)
		}
		if got["items"].([]any)[0].(M)["id"] != "s1" {
			t.Fatalf("got %v", got)
		}
	}
}

func TestSearchMessagesRejectsBadQuery(t *testing.T) {
	for query, want := range map[string]string{
		"":               "`query` is required",
		"   ":            "`query` is required",
		`"weekly report`: "`query` has an unbalanced double quote; close the phrase or drop the \"",
	} {
		cs, f := newSession(t, oneMessage)
		if msg := callErr(t, cs, "search_messages", M{"query": query}); msg != want {
			t.Fatalf("%q: %q", query, msg)
		}
		if len(f.Requests) != 0 {
			t.Fatal("called Graph")
		}
	}
}

func TestSearchMessagesNextPageFollowsToken(t *testing.T) {
	cs, f := newSession(t, oneMessage)
	tok := graph.EncodePageToken("https://graph.microsoft.com/v1.0/me/messages?%24search=%22budget%20review%22&%24top=25&%24skip=25")
	callJSON(t, cs, "search_messages", M{"query": "", "page_token": *tok})
	q := f.Last().Query
	if q.Get("$search") != `"budget review"` || q.Get("$skip") != "25" {
		t.Fatalf("query %v", q)
	}
}

func TestSearchMessagesMailboxEncoding(t *testing.T) {
	cs, f := newSession(t, oneMessage)
	callJSON(t, cs, "search_messages", M{"query": "budget review", "mailbox": "alice@example.com", "limit": 5})
	r := f.Last()
	if r.RawPath != "/users/alice%40example.com/messages" || r.Query.Get("$search") != `"budget review"` || r.Query.Get("$top") != "5" {
		t.Fatalf("%s %v", r.RawPath, r.Query)
	}
}

func TestGetMessageWithBody(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"id": "m99", "body": M{"contentType": "text", "content": "hi"}})
	})
	got := callJSON(t, cs, "get_message", M{"message_id": "m99", "include_body": true})
	if f.Last().Path != "/me/messages/m99" || got["id"] != "m99" || got["body"] != "hi" {
		t.Fatalf("%s %v", f.Last().Path, got)
	}
}

func TestListAttachments(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "a1", "name": "x.pdf", "contentType": "application/pdf", "size": 10, "isInline": false}}})
	})
	got := callJSON(t, cs, "list_attachments", M{"message_id": "m1"})
	if f.Last().Path != "/me/messages/m1/attachments" {
		t.Fatalf("path %s", f.Last().Path)
	}
	if got["items"].([]any)[0].(M)["name"] != "x.pdf" || got["next_page_token"] != nil {
		t.Fatalf("got %v", got)
	}
}

func attachment(name, ct string, data []byte) graphtest.Handler {
	return func(graphtest.Request) graphtest.Response {
		a := M{"@odata.type": "#microsoft.graph.fileAttachment", "id": "a1", "contentType": ct, "size": len(data), "contentBytes": base64.StdEncoding.EncodeToString(data)}
		if name != "" {
			a["name"] = name
		}
		return jsonOK(a)
	}
}

var pdfBytes = []byte{0, 1, 2, 3}

func TestDownloadAttachmentBase64(t *testing.T) {
	cs, f := newSession(t, attachment("x.pdf", "application/pdf", pdfBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if f.Last().Path != "/me/messages/m1/attachments/a1" {
		t.Fatalf("path %s", f.Last().Path)
	}
	if got["name"] != "x.pdf" || got["content_base64"] != "AAECAw==" || got["size_bytes"] != float64(4) {
		t.Fatalf("got %v", got)
	}
}

var pngBytes = []byte("\x89PNG\r\n\x1a\nfake")

func TestDownloadAttachmentImageBlock(t *testing.T) {
	cs, _ := newSession(t, attachment("shot.png", "image/png", pngBytes))
	text, isErr, res := callTool(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if isErr || len(res.Content) != 2 {
		t.Fatalf("err %v, %d blocks: %s", isErr, len(res.Content), text)
	}
	if !strings.Contains(text, `"name": "shot.png"`) || !strings.Contains(text, `"content_type": "image/png"`) {
		t.Fatalf("meta %s", text)
	}
	img := res.Content[1].(*mcp.ImageContent)
	if img.MIMEType != "image/png" || string(img.Data) != string(pngBytes) {
		t.Fatalf("image %+v", img)
	}
}

func TestDownloadAttachmentSaveFile(t *testing.T) {
	target := filepath.Join(t.TempDir(), "out.pdf")
	cs, _ := newSession(t, attachment("x.pdf", "application/pdf", pdfBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": target})
	if got["path"] != target || got["name"] != "x.pdf" || got["size_bytes"] != float64(4) {
		t.Fatalf("got %v", got)
	}
	if _, has := got["content_base64"]; has {
		t.Fatal("content_base64 with save_path")
	}
	if b, _ := os.ReadFile(target); string(b) != string(pdfBytes) {
		t.Fatal("file content")
	}
}

func TestDownloadAttachmentSaveDirectoryUsesName(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, attachment("report.pdf", "application/pdf", pdfBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "report.pdf") {
		t.Fatalf("got %v", got)
	}
}

func TestDownloadAttachmentSaveDirectoryUnnamed(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, attachment("", "application/octet-stream", pdfBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "attachment-a1.bin") || got["size_bytes"] != float64(4) {
		t.Fatalf("got %v", got)
	}
}

func TestDownloadAttachmentNameTraversal(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, attachment("../../evil.txt", "text/plain", pdfBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "evil.txt") {
		t.Fatalf("got %v", got)
	}
	if _, err := os.Stat(filepath.Join(dir, "evil.txt")); err != nil {
		t.Fatal(err)
	}
}

func TestDownloadItemAttachmentHasNoContent(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"@odata.type": "#microsoft.graph.itemAttachment", "id": "a1", "name": "fwd", "contentType": nil, "size": 100})
	})
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if got["content_base64"] != nil || got["name"] != "fwd" {
		t.Fatalf("got %v", got)
	}
}
