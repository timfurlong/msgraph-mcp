package tools

import (
	"crypto/sha256"
	"encoding/hex"
	"os"
	"path/filepath"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func chatMsg(id string) M {
	return M{"id": id, "from": M{"user": M{"id": "u1", "displayName": "Alice"}}, "body": M{"contentType": "html", "content": "<p>hi</p>"}}
}

func TestListChats(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "c1", "chatType": "group", "members": []any{M{"displayName": "Alice"}}, "lastMessagePreview": M{"createdDateTime": "2026-07-03T17:46:28Z"}}},
			"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$skiptoken=x"})
	})
	got := callJSON(t, cs, "list_chats", M{"limit": 10})
	r := f.Last()
	if r.Path != "/me/chats" || r.Query.Get("$top") != "10" || r.Query.Get("$expand") != "members,lastMessagePreview" {
		t.Fatalf("%s %v", r.Path, r.Query)
	}
	item := got["items"].([]any)[0].(M)
	if item["last_message_time"] != "2026-07-03T17:46:28Z" || got["next_page_token"] == nil {
		t.Fatalf("got %v", got)
	}
}

func TestListChatMessages(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(M{"value": []any{chatMsg("m1")}}) })
	got := callJSON(t, cs, "list_chat_messages", M{"chat_id": "19:abc@thread.v2", "include_body": true})
	r := f.Last()
	if r.RawPath != "/chats/19%3Aabc%40thread.v2/messages" || r.Query.Get("$top") != "25" {
		t.Fatalf("%s %v", r.RawPath, r.Query)
	}
	if got["items"].([]any)[0].(M)["body"] != "<p>hi</p>" {
		t.Fatalf("got %v", got)
	}
}

func TestListChatMessagesLimit50(t *testing.T) {
	cs, _ := newSession(t, nil)
	if msg := callErr(t, cs, "list_chat_messages", M{"chat_id": "c", "limit": 51}); msg != "limit must be an integer in 1..50 (got 51)" {
		t.Fatalf("msg %q", msg)
	}
}

func TestListJoinedTeamsCapsClientSide(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "t1", "displayName": "A"}, M{"id": "t2", "displayName": "B"}, M{"id": "t3", "displayName": "C"}}})
	})
	got := callJSON(t, cs, "list_joined_teams", M{"limit": 2})
	if f.Last().Path != "/me/joinedTeams" || len(f.Last().Query) != 0 || len(got["items"].([]any)) != 2 {
		t.Fatalf("%s %v %v", f.Last().Path, f.Last().Query, got)
	}
}

func TestListChannelsCapsClientSide(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{M{"id": "ch1", "displayName": "General", "membershipType": "standard"}, M{"id": "ch2"}}})
	})
	got := callJSON(t, cs, "list_channels", M{"team_id": "t1", "limit": 1})
	if f.Last().Path != "/teams/t1/channels" || len(f.Last().Query) != 0 || len(got["items"].([]any)) != 1 || got["items"].([]any)[0].(M)["membership_type"] != "standard" {
		t.Fatalf("%s %v", f.Last().Path, got)
	}
}

func TestListChannelMessagesAndReplies(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(M{"value": []any{chatMsg("m1")}}) })
	callJSON(t, cs, "list_channel_messages", M{"team_id": "t1", "channel_id": "ch1", "limit": 50})
	if r := f.Last(); r.Path != "/teams/t1/channels/ch1/messages" || r.Query.Get("$top") != "50" {
		t.Fatalf("%s %v", r.Path, r.Query)
	}
	callJSON(t, cs, "list_message_replies", M{"team_id": "t1", "channel_id": "ch1", "message_id": "m1", "limit": 5})
	if r := f.Last(); r.Path != "/teams/t1/channels/ch1/messages/m1/replies" || r.Query.Get("$top") != "5" {
		t.Fatalf("%s %v", r.Path, r.Query)
	}
	for _, tool := range []string{"list_channel_messages", "list_message_replies"} {
		if msg := callErr(t, cs, tool, M{"team_id": "t1", "channel_id": "ch1", "message_id": "m1", "limit": 51}); msg != "limit must be an integer in 1..50 (got 51)" {
			t.Fatalf("%s: %q", tool, msg)
		}
	}
}

func hosted(data []byte) graphtest.Handler {
	return func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Bytes: data, Header: map[string]string{"Content-Type": "application/octet-stream"}}
	}
}

func TestDownloadHostedContentImage(t *testing.T) {
	cs, f := newSession(t, hosted(pngBytes))
	text, isErr, res := callTool(t, cs, "download_hosted_content", M{"chat_id": "c1", "message_id": "m1", "hosted_content_id": "h1"})
	if isErr || len(res.Content) != 2 || f.Last().Path != "/chats/c1/messages/m1/hostedContents/h1/$value" {
		t.Fatalf("%v %s %s", isErr, text, f.Last().Path)
	}
	if img := res.Content[1].(*mcp.ImageContent); img.MIMEType != "image/png" {
		t.Fatalf("mime %s", img.MIMEType)
	}
	if text != "{\n  \"content_type\": \"image/png\",\n  \"size_bytes\": 12\n}" {
		t.Fatalf("meta %s", text)
	}
}

func TestDownloadHostedContentChannelBase64(t *testing.T) {
	cs, f := newSession(t, hosted([]byte("%PDF")))
	got := callJSON(t, cs, "download_hosted_content", M{"team_id": "t1", "channel_id": "ch1", "message_id": "m1", "hosted_content_id": "h1", "include_raw": true})
	if f.Last().Path != "/teams/t1/channels/ch1/messages/m1/hostedContents/h1/$value" {
		t.Fatalf("path %s", f.Last().Path)
	}
	if got["content_type"] != nil || got["size_bytes"] != float64(4) || got["content_base64"] != "JVBERg==" || got["raw"] == nil {
		t.Fatalf("got %v", got)
	}
}

func TestDownloadHostedContentSaveToDirectory(t *testing.T) {
	dir := t.TempDir()
	cs, _ := newSession(t, hosted(pngBytes))
	got := callJSON(t, cs, "download_hosted_content", M{"chat_id": "c1", "message_id": "m1", "hosted_content_id": "aGVsbG8=", "save_path": dir})
	sum := sha256.Sum256([]byte("aGVsbG8="))
	want := filepath.Join(dir, "hosted-content-m1-"+hex.EncodeToString(sum[:])[:8]+".png")
	if got["path"] != want || got["content_type"] != "image/png" {
		t.Fatalf("got %v, want path %s", got, want)
	}
	if b, _ := os.ReadFile(want); string(b) != string(pngBytes) {
		t.Fatal("content")
	}
}

func TestDownloadHostedContentLocationValidation(t *testing.T) {
	cs, _ := newSession(t, nil)
	if msg := callErr(t, cs, "download_hosted_content", M{"chat_id": "c", "team_id": "t", "message_id": "m", "hosted_content_id": "h"}); msg != "Provide either chat_id or (team_id and channel_id), not both." {
		t.Fatalf("msg %q", msg)
	}
	if msg := callErr(t, cs, "download_hosted_content", M{"team_id": "t", "message_id": "m", "hosted_content_id": "h"}); msg != "Provide a message location: chat_id, or both team_id and channel_id." {
		t.Fatalf("msg %q", msg)
	}
}
