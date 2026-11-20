package tools

import (
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func folderJSON(id, name string, total, children int) M {
	return M{"id": id, "displayName": name, "parentFolderId": "p", "totalItemCount": total, "unreadItemCount": 0, "childFolderCount": children}
}

func TestListFolders(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(M{"value": []any{folderJSON("f1", "Inbox", 3, 0)}, "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?$skip=10"})
	})
	got := callJSON(t, cs, "list_folders", nil)
	if f.Last().Path != "/me/mailFolders" || got["items"].([]any)[0].(M)["display_name"] != "Inbox" || got["next_page_token"] != nil {
		t.Fatalf("%s %v", f.Last().Path, got)
	}
}

func TestMoveMessage(t *testing.T) {
	for _, dest := range []string{"AAMkFolder", "archive", "deleteditems"} {
		cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
			return graphtest.Response{Status: 201, JSON: M{"id": "m1"}}
		})
		got := callJSON(t, cs, "move_message", M{"message_id": "m1", "destination": dest})
		r := f.Last()
		if r.Method != "POST" || r.Path != "/me/messages/m1/move" || r.Body["destinationId"] != dest || got["id"] != "m1" {
			t.Fatalf("%s %s %v %v", r.Method, r.Path, r.Body, got)
		}
	}
	cs, _ := newSession(t, nil)
	if msg := callErr(t, cs, "move_message", M{"message_id": "m1", "destination": ""}); msg != "`destination` is required" {
		t.Fatalf("msg %q", msg)
	}
}

func TestCreateFolder(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Status: 201, JSON: folderJSON("new-id", "Test notifications", 0, 0)}
	})
	got := callJSON(t, cs, "create_folder", M{"display_name": "Test notifications"})
	r := f.Last()
	if r.Method != "POST" || r.Path != "/me/mailFolders" || r.Body["displayName"] != "Test notifications" || got["id"] != "new-id" {
		t.Fatalf("%s %s %v %v", r.Method, r.Path, r.Body, got)
	}
	callJSON(t, cs, "create_folder", M{"display_name": "Sub", "parent_folder_id": "parent-id"})
	if r := f.Last(); r.Path != "/me/mailFolders/parent-id/childFolders" || r.Body["displayName"] != "Sub" {
		t.Fatalf("%s %v", r.Path, r.Body)
	}
	for _, name := range []string{"", "   "} {
		if msg := callErr(t, cs, "create_folder", M{"display_name": name}); msg != "`display_name` is required" {
			t.Fatalf("msg %q", msg)
		}
	}
}

func TestUpdateFolder(t *testing.T) {
	cs, f := newSession(t, func(r graphtest.Request) graphtest.Response {
		name := "Renamed"
		if r.Body["displayName"] == nil {
			name = "Old"
		}
		return jsonOK(folderJSON("fid", name, 0, 0))
	})
	got := callJSON(t, cs, "update_folder", M{"folder_id": "fid", "display_name": "Renamed"})
	if len(f.Requests) != 1 || f.Last().Method != "PATCH" || f.Last().Path != "/me/mailFolders/fid" || got["display_name"] != "Renamed" {
		t.Fatalf("%v %v", f.Requests, got)
	}
	callJSON(t, cs, "update_folder", M{"folder_id": "fid", "parent_folder_id": "new-parent"})
	if r := f.Last(); r.Method != "POST" || r.Path != "/me/mailFolders/fid/move" || r.Body["destinationId"] != "new-parent" || len(f.Requests) != 2 {
		t.Fatalf("%s %s %v", r.Method, r.Path, r.Body)
	}
	got = callJSON(t, cs, "update_folder", M{"folder_id": "fid", "display_name": "Renamed", "parent_folder_id": "np"})
	if len(f.Requests) != 4 || f.Requests[2].Method != "PATCH" || f.Requests[3].Path != "/me/mailFolders/fid/move" {
		t.Fatalf("requests %v", f.Requests)
	}
	for args, want := range map[string]M{
		"`folder_id` is required": {"folder_id": "", "display_name": "x"},
		"At least one of `display_name` or `parent_folder_id` is required": {"folder_id": "fid"},
		"`display_name` cannot be empty":                                   {"folder_id": "fid", "display_name": "  "},
	} {
		if msg := callErr(t, cs, "update_folder", want); msg != args {
			t.Fatalf("msg %q, want %q", msg, args)
		}
	}
}

func TestDeleteFolder(t *testing.T) {
	cs, f := newSession(t, func(r graphtest.Request) graphtest.Response {
		if r.Method == "DELETE" {
			return graphtest.Response{Status: 204}
		}
		return jsonOK(folderJSON("fid", "Empty", 0, 0))
	})
	got := callJSON(t, cs, "delete_folder", M{"folder_id": "fid"})
	if got["deleted"] != true || got["id"] != "fid" || f.Requests[0].Method != "GET" || f.Last().Method != "DELETE" || f.Last().Path != "/me/mailFolders/fid" {
		t.Fatalf("%v %v", got, f.Requests)
	}
	for _, c := range []struct{ total, children int }{{2, 0}, {0, 1}} {
		cs, f := newSession(t, func(graphtest.Request) graphtest.Response {
			return jsonOK(folderJSON("fid", "Full", c.total, c.children))
		})
		msg := callErr(t, cs, "delete_folder", M{"folder_id": "fid"})
		want := "Refusing to delete non-empty folder 'fid': contains " + itoa(c.total) + " message(s) and " + itoa(c.children) + " child folder(s). Empty it first or move the contents elsewhere."
		if msg != want || len(f.Requests) != 1 {
			t.Fatalf("msg %q", msg)
		}
	}
	if msg := callErr(t, cs, "delete_folder", M{"folder_id": ""}); msg != "`folder_id` is required" {
		t.Fatalf("msg %q", msg)
	}
}

func TestMailActions(t *testing.T) {
	cases := []struct {
		tool   string
		method string
		path   string
		body   string
	}{
		{"archive_message", "POST", "/me/messages/m1/move", `{"destinationId":"archive"}`},
		{"mark_read", "PATCH", "/me/messages/m1", `{"isRead":true}`},
		{"mark_unread", "PATCH", "/me/messages/m1", `{"isRead":false}`},
		{"flag_message", "PATCH", "/me/messages/m1", `{"flag":{"flagStatus":"flagged"}}`},
		{"unflag_message", "PATCH", "/me/messages/m1", `{"flag":{"flagStatus":"notFlagged"}}`},
	}
	for _, c := range cases {
		cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(M{"id": "m1", "isRead": true}) })
		got := callJSON(t, cs, c.tool, M{"message_id": "m1", "include_raw": true})
		r := f.Last()
		if r.Method != c.method || r.Path != c.path || string(r.RawBody) != c.body || got["id"] != "m1" || got["raw"] == nil {
			t.Fatalf("%s: %s %s %s %v", c.tool, r.Method, r.Path, r.RawBody, got)
		}
	}
}
