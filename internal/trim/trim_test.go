package trim

import (
	"encoding/json"
	"os"
	"reflect"
	"strings"
	"testing"
	"unicode/utf8"
)

func load(t *testing.T, name string) M {
	t.Helper()
	b, err := os.ReadFile("testdata/" + name)
	if err != nil {
		t.Fatal(err)
	}
	return parse(t, string(b))
}

// parse decodes JSON the way the Graph client does (UseNumber).
func parse(t *testing.T, s string) M {
	t.Helper()
	dec := json.NewDecoder(strings.NewReader(s))
	dec.UseNumber()
	var m M
	if err := dec.Decode(&m); err != nil {
		t.Fatal(err)
	}
	return m
}

// field marshals got[key] and compares it to the JSON literal want.
func field(t *testing.T, got M, key, want string) {
	t.Helper()
	v, ok := got[key]
	if !ok {
		t.Fatalf("key %q missing in %v", key, got)
	}
	b, _ := json.Marshal(v)
	var a, w any
	_ = json.Unmarshal(b, &a)
	if err := json.Unmarshal([]byte(want), &w); err != nil {
		t.Fatalf("bad want %q: %v", want, err)
	}
	if !reflect.DeepEqual(a, w) {
		t.Fatalf("%s = %s, want %s", key, b, want)
	}
}

func equalJSON(t *testing.T, got M, want string) {
	t.Helper()
	b, _ := json.Marshal(got)
	var a, w any
	_ = json.Unmarshal(b, &a)
	_ = json.Unmarshal([]byte(want), &w)
	if !reflect.DeepEqual(a, w) {
		t.Fatalf("got %s\nwant %s", b, want)
	}
}

func absent(t *testing.T, got M, key string) {
	t.Helper()
	if _, ok := got[key]; ok {
		t.Fatalf("key %q present: %v", key, got[key])
	}
}

func TestMessageMinimal(t *testing.T) {
	m := Message(load(t, "message_minimal.json"), false, false)
	field(t, m, "id", `"AAMkAGI-msg-id"`)
	field(t, m, "subject", `"Hello"`)
	field(t, m, "from", `{"name":"Alice","address":"alice@example.com"}`)
	field(t, m, "to", `[{"name":"Bob","address":"bob@example.com"}]`)
	field(t, m, "cc", `[]`)
	field(t, m, "received", `"2026-05-19T14:00:00Z"`)
	field(t, m, "sent", `"2026-05-19T13:59:00Z"`)
	field(t, m, "snippet", `"Hello there"`)
	field(t, m, "is_read", `false`)
	field(t, m, "flag", `"notFlagged"`)
	field(t, m, "categories", `[]`)
	field(t, m, "conversation_id", `"AAQk-conv"`)
	absent(t, m, "body")
	absent(t, m, "raw")
	absent(t, m, "bcc")
}

func TestMessageWithBodyAndRaw(t *testing.T) {
	raw := load(t, "message_minimal.json")
	m := Message(raw, true, true)
	field(t, m, "body", `"Hello there\nFull body."`)
	field(t, m, "body_type", `"text"`)
	if !reflect.DeepEqual(m["raw"], raw) {
		t.Fatal("raw is not the Graph object")
	}
}

func TestMessageMissingOptionalFields(t *testing.T) {
	m := Message(M{"id": "x", "subject": "y", "isRead": true, "isDraft": false, "hasAttachments": false}, false, false)
	field(t, m, "from", `null`)
	field(t, m, "to", `[]`)
	field(t, m, "cc", `[]`)
	field(t, m, "received", `null`)
	field(t, m, "flag", `"notFlagged"`)
	field(t, m, "importance", `"normal"`)
	field(t, m, "body_type", `"text"`)
	field(t, m, "parent_folder_id", `null`)
	field(t, m, "inference_classification", `null`)
	field(t, m, "is_read", `true`)
}

func TestMessageFromFallsBackToSender(t *testing.T) {
	m := Message(M{"sender": M{"emailAddress": M{"name": "S", "address": "s@example.com"}}}, false, false)
	field(t, m, "from", `{"name":"S","address":"s@example.com"}`)
}

func TestMessageParentFolderAndInference(t *testing.T) {
	m := Message(M{"id": "x", "parentFolderId": "AAMkAGI-folder-id", "inferenceClassification": "focused"}, false, false)
	field(t, m, "parent_folder_id", `"AAMkAGI-folder-id"`)
	field(t, m, "inference_classification", `"focused"`)
}

func TestRawIsGraphJSON(t *testing.T) {
	raw := parse(t, `{"id":"x","@odata.etag":"W/\"abc\"","subject":"s"}`)
	m := Message(raw, false, true)
	if !reflect.DeepEqual(m["raw"], raw) {
		t.Fatalf("raw = %v", m["raw"])
	}
}

func TestEventMinimal(t *testing.T) {
	e := Event(load(t, "event_minimal.json"), false, false)
	field(t, e, "subject", `"Standup"`)
	field(t, e, "organizer", `{"name":"Alice","address":"alice@example.com"}`)
	field(t, e, "start", `{"date_time":"2026-05-19T15:00:00","time_zone":"America/Los_Angeles"}`)
	field(t, e, "location", `"Zoom"`)
	field(t, e, "is_online_meeting", `true`)
	field(t, e, "online_meeting_url", `"https://zoom.us/j/123"`)
	field(t, e, "attendees", `[{"name":"Bob","address":"bob@example.com","type":"required","response":"accepted"}]`)
	field(t, e, "is_reminder_on", `true`)
	field(t, e, "reminder_minutes_before_start", `15`)
	field(t, e, "recurrence", `null`)
	absent(t, e, "body")
}

func TestEventWithBodyAndRaw(t *testing.T) {
	raw := load(t, "event_minimal.json")
	e := Event(raw, true, true)
	field(t, e, "body", `"<p>Daily standup</p>"`)
	if !reflect.DeepEqual(e["raw"], raw) {
		t.Fatal("raw mismatch")
	}
}

func TestEventOfflineHasNoMeetingURL(t *testing.T) {
	e := Event(M{"isOnlineMeeting": false, "onlineMeeting": M{"joinUrl": "u"}}, false, false)
	absent(t, e, "online_meeting_url")
	field(t, e, "attendees", `[]`)
}

func TestEventRecurrenceIsJSON(t *testing.T) {
	raw := parse(t, `{"recurrence":{"pattern":{"type":"weekly","interval":1,"daysOfWeek":["monday"]},"range":{"type":"noEnd","startDate":"2026-01-05"}}}`)
	e := Event(raw, false, false)
	field(t, e, "recurrence", `{"pattern":{"type":"weekly","interval":1,"daysOfWeek":["monday"]},"range":{"type":"noEnd","startDate":"2026-01-05"}}`)
}

func TestFolder(t *testing.T) {
	equalJSON(t, Folder(parse(t, `{"id":"f1","displayName":"Inbox","parentFolderId":"p1","totalItemCount":42,"unreadItemCount":3,"childFolderCount":0}`), false),
		`{"id":"f1","display_name":"Inbox","parent_folder_id":"p1","total_item_count":42,"unread_item_count":3,"child_folder_count":0}`)
}

func TestAttachmentItem(t *testing.T) {
	equalJSON(t, AttachmentItem(parse(t, `{"id":"a1","name":"doc.pdf","contentType":"application/pdf","size":12345,"isInline":false}`), false),
		`{"id":"a1","name":"doc.pdf","content_type":"application/pdf","size_bytes":12345,"is_inline":false}`)
}

func TestAttachmentDownload(t *testing.T) {
	equalJSON(t, AttachmentDownload(parse(t, `{"name":"a.txt","contentType":"text/plain","size":3,"contentBytes":"YWJj"}`), false),
		`{"name":"a.txt","content_type":"text/plain","size_bytes":3,"content_base64":"YWJj"}`)
}

func TestUser(t *testing.T) {
	equalJSON(t, User(parse(t, `{"id":"u1","displayName":"Alice","userPrincipalName":"alice@example.com","mail":"alice@example.com","jobTitle":"Engineer"}`), false),
		`{"id":"u1","display_name":"Alice","user_principal_name":"alice@example.com","mail":"alice@example.com","job_title":"Engineer"}`)
}

func TestCalendar(t *testing.T) {
	equalJSON(t, Calendar(parse(t, `{"id":"c1","name":"Calendar","owner":{"name":"Alice","address":"alice@example.com"},"canEdit":true,"isDefaultCalendar":true}`), false),
		`{"id":"c1","name":"Calendar","owner":{"name":"Alice","address":"alice@example.com"},"can_edit":true,"is_default_calendar":true}`)
	equalJSON(t, Calendar(M{"id": "c2"}, false), `{"id":"c2","name":null,"owner":null,"can_edit":false,"is_default_calendar":false}`)
}

func TestRuleMinimal(t *testing.T) {
	r := Rule(parse(t, `{"id":"rule-1","displayName":"Test notifications","sequence":1,"isEnabled":true,"hasError":false,"isReadOnly":false,
		"conditions":{"senderContains":["example.com"],"subjectContains":null,"bodyContains":[],"fromAddresses":[]},
		"actions":{"moveToFolder":"AAMkFolderId","markAsRead":null,"delete":null,"stopProcessingRules":true,"forwardTo":[]}}`), false)
	field(t, r, "sequence", `1`)
	field(t, r, "is_enabled", `true`)
	field(t, r, "conditions", `{"sender_contains":["example.com"],"subject_contains":null,"body_contains":null,"body_or_subject_contains":null,"from_addresses":[],"has_attachments":null,"categories":null}`)
	field(t, r, "actions", `{"move_to_folder":"AAMkFolderId","copy_to_folder":null,"delete":null,"mark_as_read":null,"mark_importance":null,"assign_categories":null,"forward_to":[],"stop_processing_rules":true}`)
	absent(t, r, "raw")
	absent(t, r, "exceptions")
}

func TestRuleMissingOptionalFields(t *testing.T) {
	raw := M{"id": "r", "displayName": "n"}
	r := Rule(raw, true)
	field(t, r, "sequence", `null`)
	field(t, r, "is_enabled", `false`)
	field(t, r, "has_error", `false`)
	field(t, r, "conditions", `null`)
	field(t, r, "actions", `null`)
	if !reflect.DeepEqual(r["raw"], raw) {
		t.Fatal("raw mismatch")
	}
}

func TestRuleFromAddresses(t *testing.T) {
	r := Rule(parse(t, `{"id":"rule-3","conditions":{"fromAddresses":[{"emailAddress":{"name":"Alice","address":"alice@example.com"}},{"emailAddress":{"name":"Bob","address":"bob@example.com"}}]},
		"actions":{"forwardTo":[{"emailAddress":{"name":"Forwardee","address":"fwd@example.com"}}]}}`), false)
	field(t, r["conditions"].(M), "from_addresses", `[{"name":"Alice","address":"alice@example.com"},{"name":"Bob","address":"bob@example.com"}]`)
	field(t, r["actions"].(M), "forward_to", `[{"name":"Forwardee","address":"fwd@example.com"}]`)
}

func chatMessage(t *testing.T, over string) M {
	base := parse(t, `{"id":"m1","messageType":"message","createdDateTime":"2026-07-01T10:00:00Z","lastModifiedDateTime":"2026-07-01T10:00:00Z",
		"deletedDateTime":null,"importance":"normal","subject":null,
		"from":{"user":{"id":"u1","displayName":"Alice"},"application":null},
		"body":{"contentType":"html","content":"<p>Hello <b>there</b> &amp; welcome</p>"},
		"attachments":[{"id":"a1","name":"f.docx","contentType":"reference","contentUrl":"https://example.com/f.docx"}],
		"mentions":[{"id":0,"mentionText":"Bob"},{"id":1,"mentionText":""}],
		"reactions":[{"reactionType":"like"},{"reactionType":"like"},{"reactionType":"heart"}],
		"webUrl":"https://teams.example/msg/1"}`)
	if over != "" {
		for k, v := range parse(t, over) {
			base[k] = v
		}
	}
	return base
}

func TestChatMessageSnippetStripsHTML(t *testing.T) {
	m := ChatMessage(chatMessage(t, ""), false, false)
	field(t, m, "snippet", `"Hello there & welcome"`)
	field(t, m, "from", `"Alice"`)
	field(t, m, "from_id", `"u1"`)
	field(t, m, "body_type", `"html"`)
	field(t, m, "deleted", `false`)
	absent(t, m, "body")
}

func TestChatMessageAppSenderHasNoFrom(t *testing.T) {
	m := ChatMessage(chatMessage(t, `{"from":{"user":null,"application":{"id":"app","displayName":"Bot"}}}`), false, false)
	field(t, m, "from", `null`)
	field(t, m, "from_id", `null`)
}

func TestChatMessageReactionsMentionsAttachments(t *testing.T) {
	m := ChatMessage(chatMessage(t, ""), false, false)
	field(t, m, "reactions", `{"like":2,"heart":1}`)
	field(t, m, "mentions", `["Bob"]`)
	field(t, m, "attachments", `[{"id":"a1","name":"f.docx","content_type":"reference","content_url":"https://example.com/f.docx"}]`)
}

func TestChatMessageIncludeBody(t *testing.T) {
	m := ChatMessage(chatMessage(t, ""), true, false)
	field(t, m, "body", `"<p>Hello <b>there</b> &amp; welcome</p>"`)
}

func TestChatMessageHostedRefs(t *testing.T) {
	m := ChatMessage(chatMessage(t, `{"body":{"contentType":"html","content":"<div><img src=\"https://graph.microsoft.com/v1.0/chats/c1/messages/m1/hostedContents/abc123/$value\"><img src=\"https://graph.microsoft.com/v1.0/chats/c1/messages/m1/hostedContents/abc123/$value\"><img src='x/hostedContents/def/$value'></div>"}}`), false, false)
	field(t, m, "hosted_content_refs", `[{"hosted_content_id":"abc123"},{"hosted_content_id":"def"}]`)
}

const card = `{"type": "AdaptiveCard", "body": [{"type": "TextBlock", "text": "Access issue detected"}, {"type": "Container", "items": [{"type": "TextBlock", "text": "User bob@example.com locked out"}]}, {"text": "   "}]}`

func cardMessage(t *testing.T) M {
	cardJSON, _ := json.Marshal(card)
	return chatMessage(t, `{"body":{"contentType":"html","content":"<attachment id=\"a1\"></attachment>"},
		"attachments":[{"id":"a1","contentType":"application/vnd.microsoft.card.adaptive","contentUrl":null,"name":null,"content":`+string(cardJSON)+`}]}`)
}

func TestChatMessageCardSnippetFallback(t *testing.T) {
	m := ChatMessage(cardMessage(t), false, false)
	field(t, m, "snippet", `"Access issue detected User bob@example.com locked out"`)
	if _, has := m["attachments"].([]any)[0].(M)["content"]; has {
		t.Fatal("attachment content without include_body")
	}
}

func TestChatMessageCardContentWithIncludeBody(t *testing.T) {
	m := ChatMessage(cardMessage(t), true, false)
	if c, _ := m["attachments"].([]any)[0].(M)["content"].(string); !strings.Contains(c, `"Access issue detected"`) {
		t.Fatalf("content = %q", c)
	}
}

func TestCardTextKeepsDocumentOrder(t *testing.T) {
	// Keys deliberately out of alphabetical order.
	got := CardText([]any{M{"contentType": "application/vnd.microsoft.card.adaptive", "content": `{"z":{"text":"first"},"a":{"text":"second"},"m":[{"text":"third"}]}`}})
	if got == nil || *got != "first second third" {
		t.Fatalf("got %v", got)
	}
}

func TestChatMessageDeletedFlag(t *testing.T) {
	m := ChatMessage(chatMessage(t, `{"deletedDateTime":"2026-07-02T00:00:00Z"}`), false, false)
	field(t, m, "deleted", `true`)
}

func TestChatMessageSnippetTruncates(t *testing.T) {
	m := ChatMessage(chatMessage(t, `{"body":{"contentType":"text","content":"`+strings.Repeat("x", 400)+`"}}`), false, false)
	s := m["snippet"].(string)
	if !strings.HasSuffix(s, "...") || len(s) != 283 {
		t.Fatalf("snippet len %d", len(s))
	}
}

func TestSnippetRuneSafe(t *testing.T) {
	text := strings.Repeat("a", 279) + "😀é" + strings.Repeat("b", 50)
	got := Snippet(text, "text")
	if got == nil {
		t.Fatal("nil snippet")
	}
	if !utf8.ValidString(*got) {
		t.Fatal("invalid UTF-8")
	}
	want := strings.Repeat("a", 279) + "😀..."
	if *got != want {
		t.Fatalf("got %q", *got)
	}
}

func TestSnippetUnicodeWhitespaceAndEntities(t *testing.T) {
	got := Snippet("<p>a&nbsp;&nbsp;b</p>\n\n<p>c</p>", "HTML")
	if got == nil || *got != "a b c" {
		t.Fatalf("got %v", got)
	}
	if Snippet("", "html") != nil || Snippet("<p> </p>", "html") != nil {
		t.Fatal("empty snippet should be nil")
	}
	// Plain text keeps angle brackets.
	if got := Snippet("a <b> c", "text"); *got != "a <b> c" {
		t.Fatalf("got %q", *got)
	}
}

func TestChatMessageAllMissing(t *testing.T) {
	m := ChatMessage(M{}, false, false)
	equalJSON(t, m, `{"id":null,"message_type":"message","from":null,"from_id":null,"created":null,"last_modified":null,"deleted":false,"importance":"normal","subject":null,"snippet":null,"body_type":"text","attachments":[],"mentions":[],"hosted_content_refs":[],"reactions":{},"web_url":null}`)
}

func TestChat(t *testing.T) {
	c := Chat(parse(t, `{"id":"c1","chatType":"group","topic":"Launch","members":[{"displayName":"Alice"},{"displayName":null},{"displayName":"Bob"}],
		"lastUpdatedDateTime":"2026-07-01T09:00:00Z","lastMessagePreview":{"createdDateTime":"2026-07-03T17:46:28Z"},"webUrl":"https://teams.example/chat/c1"}`), false)
	equalJSON(t, c, `{"id":"c1","chat_type":"group","topic":"Launch","members":["Alice","Bob"],"last_message_time":"2026-07-03T17:46:28Z","metadata_updated":"2026-07-01T09:00:00Z","web_url":"https://teams.example/chat/c1"}`)
}

func TestChatMissingPreview(t *testing.T) {
	c := Chat(parse(t, `{"id":"c1","chatType":"oneOnOne","lastUpdatedDateTime":"2025-12-22T17:46:20Z"}`), false)
	field(t, c, "last_message_time", `null`)
	field(t, c, "members", `[]`)
}

func TestTeamAndChannel(t *testing.T) {
	equalJSON(t, Team(M{"id": "t1", "displayName": "Eng", "description": "d"}, false), `{"id":"t1","display_name":"Eng","description":"d"}`)
	equalJSON(t, Channel(M{"id": "ch1", "displayName": "General", "description": nil, "membershipType": "standard", "webUrl": "u"}, false),
		`{"id":"ch1","display_name":"General","description":null,"membership_type":"standard","web_url":"u"}`)
}

func TestHostedContentDownload(t *testing.T) {
	equalJSON(t, HostedDownload(M{"contentType": "image/png", "size": 4, "contentBytes": "AAECAw=="}, false), `{"content_type":"image/png","size_bytes":4,"content_base64":"AAECAw=="}`)
}

