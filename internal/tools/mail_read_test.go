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

// itemAttachment scripts an item attachment: the metadata GET (with the
// nested item expanded) and the /$value GET that returns its raw content.
func itemAttachment(name string, item M, value []byte) graphtest.Handler {
	return func(r graphtest.Request) graphtest.Response {
		if strings.HasSuffix(r.Path, "/$value") {
			return graphtest.Response{Bytes: value}
		}
		a := M{"@odata.type": "#microsoft.graph.itemAttachment", "id": "a1", "name": name, "contentType": "message/rfc822", "size": 9000, "isInline": false}
		if r.Query.Get("$expand") == "microsoft.graph.itemattachment/item" {
			a["item"] = item
		}
		return jsonOK(a)
	}
}

var emlBytes = []byte("From: Alice <alice@example.com>\r\nSubject: Quarterly report\r\n\r\nHello\r\n")

var attachedMessage = M{
	"@odata.type":  "#microsoft.graph.message",
	"id":           "",
	"subject":      "Quarterly report",
	"from":         M{"emailAddress": M{"name": "Alice", "address": "alice@example.com"}},
	"toRecipients": []any{M{"emailAddress": M{"name": "Bob", "address": "bob@example.com"}}},
	"sentDateTime": "2026-03-20T23:15:48Z",
	"body":         M{"contentType": "html", "content": "<p>Hello</p>"},
	"bodyPreview":  "Hello",
}

func TestDownloadItemAttachmentMessage(t *testing.T) {
	cs, f := newSession(t, itemAttachment("Quarterly report", attachedMessage, emlBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if len(f.Requests) != 2 {
		t.Fatalf("%d requests", len(f.Requests))
	}
	meta, value := f.Requests[0], f.Requests[1]
	if meta.Path != "/me/messages/m1/attachments/a1" || meta.Query.Get("$expand") != "microsoft.graph.itemattachment/item" {
		t.Fatalf("meta request %s %v", meta.Path, meta.Query)
	}
	if value.Path != "/me/messages/m1/attachments/a1/$value" {
		t.Fatalf("value path %s", value.Path)
	}
	if got["kind"] != "item" || got["item_type"] != "message" || got["content_type"] != "message/rfc822" {
		t.Fatalf("got %v", got)
	}
	if got["content_text"] != string(emlBytes) || got["size_bytes"] != float64(len(emlBytes)) {
		t.Fatalf("content %v", got)
	}
	if _, has := got["content_base64"]; has {
		t.Fatal("content_base64 alongside content_text")
	}
	item := got["item"].(M)
	if item["subject"] != "Quarterly report" || item["from"].(M)["address"] != "alice@example.com" || item["body"] != "<p>Hello</p>" {
		t.Fatalf("item %v", item)
	}
	if item["to"].([]any)[0].(M)["address"] != "bob@example.com" || item["sent"] != "2026-03-20T23:15:48Z" {
		t.Fatalf("item %v", item)
	}
}

func TestDownloadItemAttachmentSaveEml(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, itemAttachment("Quarterly report", attachedMessage, emlBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	want := filepath.Join(dir, "Quarterly report.eml")
	if got["path"] != want || got["name"] != "Quarterly report" || got["content_type"] != "message/rfc822" || got["size_bytes"] != float64(len(emlBytes)) {
		t.Fatalf("got %v", got)
	}
	if len(got) != 4 {
		t.Fatalf("want only path, name, content_type, size_bytes; got %v", got)
	}
	if b, _ := os.ReadFile(want); string(b) != string(emlBytes) {
		t.Fatal("file content")
	}
}

func TestDownloadItemAttachmentSaveFilePath(t *testing.T) {
	target := filepath.Join(t.TempDir(), "sub", "fwd.eml")
	cs, _ := newSession(t, itemAttachment("Quarterly report", attachedMessage, emlBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": target})
	if got["path"] != target {
		t.Fatalf("got %v", got)
	}
	if b, _ := os.ReadFile(target); string(b) != string(emlBytes) {
		t.Fatal("file content")
	}
}

func TestDownloadItemAttachmentNameWithSlashes(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, itemAttachment("REVIEW // plan", attachedMessage, emlBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "REVIEW -- plan.eml") {
		t.Fatalf("got %v", got)
	}
}

func TestDownloadItemAttachmentUnnamed(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, itemAttachment("", attachedMessage, emlBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "attachment-a1.eml") {
		t.Fatalf("got %v", got)
	}
}

var icsBytes = []byte("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nSUMMARY:Planning\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")

var attachedEvent = M{
	"@odata.type": "#microsoft.graph.event",
	"subject":     "Planning",
	"organizer":   M{"emailAddress": M{"name": "Alice", "address": "alice@example.com"}},
	"start":       M{"dateTime": "2026-12-01T15:00:00.0000000", "timeZone": "UTC"},
	"end":         M{"dateTime": "2026-12-01T16:00:00.0000000", "timeZone": "UTC"},
	"location":    M{"displayName": "Room 1"},
}

func TestDownloadItemAttachmentEvent(t *testing.T) {
	cs, _ := newSession(t, itemAttachment("Planning", attachedEvent, icsBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if got["item_type"] != "event" || got["content_type"] != "text/calendar" || got["content_text"] != string(icsBytes) {
		t.Fatalf("got %v", got)
	}
	item := got["item"].(M)
	if item["subject"] != "Planning" || item["location"] != "Room 1" || item["start"].(M)["date_time"] != "2026-12-01T15:00:00.0000000" {
		t.Fatalf("item %v", item)
	}
}

func TestDownloadItemAttachmentEventSavesIcs(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, itemAttachment("Planning", attachedEvent, icsBytes))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "save_path": dir})
	if got["path"] != filepath.Join(dir, "Planning.ics") || got["content_type"] != "text/calendar" {
		t.Fatalf("got %v", got)
	}
}

func TestDownloadItemAttachmentBinaryContentIsBase64(t *testing.T) {
	raw := []byte{0xff, 0xfe, 0x00, 0x01}
	cs, _ := newSession(t, itemAttachment("odd", attachedMessage, raw))
	got := callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if got["content_base64"] != base64.StdEncoding.EncodeToString(raw) {
		t.Fatalf("got %v", got)
	}
	if _, has := got["content_text"]; has {
		t.Fatal("content_text for non-UTF-8 content")
	}
}

func TestDownloadItemAttachmentSharedMailbox(t *testing.T) {
	cs, f := newSession(t, itemAttachment("Quarterly report", attachedMessage, emlBytes))
	callJSON(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1", "mailbox": "alice@example.com"})
	if f.Requests[0].RawPath != "/users/alice%40example.com/messages/m1/attachments/a1" {
		t.Fatalf("meta path %s", f.Requests[0].RawPath)
	}
	if f.Requests[1].RawPath != "/users/alice%40example.com/messages/m1/attachments/a1/$value" {
		t.Fatalf("value path %s", f.Requests[1].RawPath)
	}
}

func TestDownloadReferenceAttachmentErrors(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"@odata.type": "#microsoft.graph.referenceAttachment", "id": "a1", "name": "Plan.docx", "contentType": nil, "size": 300})
	})
	text := callErr(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if !strings.Contains(text, "referenceAttachment") || !strings.Contains(text, "Plan.docx") {
		t.Fatalf("error %s", text)
	}
	if len(f.Requests) != 1 {
		t.Fatalf("%d requests, want only the metadata GET", len(f.Requests))
	}
}

func TestDownloadUnknownAttachmentKindErrors(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"@odata.type": "#microsoft.graph.somethingNew", "id": "a1", "name": "x"})
	})
	text := callErr(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if !strings.Contains(text, "somethingNew") {
		t.Fatalf("error %s", text)
	}
}

func TestDownloadFileAttachmentWithoutBytesErrors(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"@odata.type": "#microsoft.graph.fileAttachment", "id": "a1", "name": "x.pdf"})
	})
	text := callErr(t, cs, "download_attachment", M{"message_id": "m1", "attachment_id": "a1"})
	if !strings.Contains(text, "fileAttachment") {
		t.Fatalf("error %s", text)
	}
}

func TestListAttachmentsKind(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{
			M{"@odata.type": "#microsoft.graph.fileAttachment", "id": "a1", "name": "x.pdf"},
			M{"@odata.type": "#microsoft.graph.itemAttachment", "id": "a2", "name": "fwd"},
			M{"@odata.type": "#microsoft.graph.referenceAttachment", "id": "a3", "name": "Plan.docx"},
		}})
	})
	got := callJSON(t, cs, "list_attachments", M{"message_id": "m1"})
	items := got["items"].([]any)
	for i, want := range []string{"file", "item", "reference"} {
		if items[i].(M)["kind"] != want {
			t.Fatalf("item %d kind %v, want %s", i, items[i].(M)["kind"], want)
		}
	}
}
