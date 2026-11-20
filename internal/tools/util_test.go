package tools

import (
	"context"
	"errors"
	"strings"
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

var me = map[string]any{"id": "u1", "displayName": "Alice", "userPrincipalName": "alice@example.com", "mail": "alice@example.com", "jobTitle": "Engineer", "extra": "x"}

func TestWhoami(t *testing.T) {
	cs, f := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(me) })
	got := callJSON(t, cs, "whoami", nil)
	if got["display_name"] != "Alice" || got["mail"] != "alice@example.com" || got["job_title"] != "Engineer" {
		t.Fatalf("got %v", got)
	}
	if _, has := got["raw"]; has {
		t.Fatal("raw without include_raw")
	}
	if r := f.Last(); r.Method != "GET" || r.Path != "/me" {
		t.Fatalf("%s %s", r.Method, r.Path)
	}
}

func TestWhoamiIncludeRaw(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(me) })
	got := callJSON(t, cs, "whoami", map[string]any{"include_raw": true})
	raw, _ := got["raw"].(map[string]any)
	if raw["extra"] != "x" {
		t.Fatalf("raw = %v", got["raw"])
	}
}

func TestOutputIsIndentedAndUnescaped(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return jsonOK(map[string]any{"id": "u1", "displayName": "<Alice & Bob>"})
	})
	text, _, _ := callTool(t, cs, "whoami", nil)
	if !strings.Contains(text, "\n  \"") || !strings.Contains(text, "<Alice & Bob>") {
		t.Fatalf("text = %s", text)
	}
}

func TestDefaultsAndExtraArgs(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response { return jsonOK(me) })
	got := callJSON(t, cs, "whoami", map[string]any{"unknown_arg": 1})
	if got["id"] != "u1" {
		t.Fatalf("got %v", got)
	}
	msg := callErr(t, cs, "whoami", map[string]any{"include_raw": "yes"})
	if msg == "" {
		t.Fatal("empty validation error")
	}
}

func TestErrorsReachCaller(t *testing.T) {
	cs, _ := newSession(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Error(403, "ErrorAccessDenied", "Access is denied.")
	})
	if msg := callErr(t, cs, "whoami", nil); msg != "Graph API 403: ErrorAccessDenied — Access is denied." {
		t.Fatalf("msg = %q", msg)
	}

	g, _ := graphtest.New(t, func(graphtest.Request) graphtest.Response { return jsonOK(me) })
	hint := "Not authenticated. Run `msgraph-mcp login` in a terminal to sign in."
	g.Tokens = errTokens{errors.New(hint)}
	cs2 := sessionFor(t, g)
	if msg := callErr(t, cs2, "whoami", nil); msg != hint {
		t.Fatalf("msg = %q", msg)
	}
}

type errTokens struct{ err error }

func (e errTokens) Token(context.Context) (string, error) { return "", e.err }
