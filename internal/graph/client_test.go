package graph_test

import (
	"context"
	"errors"
	"net/http"
	"testing"
	"time"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

func ptr(s string) *string { return &s }

func TestRootMeAndUsers(t *testing.T) {
	if got := graph.Root(nil); got != "/me" {
		t.Fatalf("Root(nil) = %q", got)
	}
	if got := graph.Root(ptr("shared@example.com")); got != "/users/shared%40example.com" {
		t.Fatalf("Root = %q", got)
	}
	// An empty string counts as set, as in Python.
	if got := graph.Root(ptr("")); got != "/users/" {
		t.Fatalf("Root(\"\") = %q", got)
	}
}

func TestRootEscapesMailbox(t *testing.T) {
	if got := graph.Root(ptr("a+b@example.com")); got != "/users/a%2Bb%40example.com" {
		t.Fatalf("Root = %q", got)
	}
	if got := graph.BatchRoot(ptr("a+b@example.com")); got != "/users/a+b@example.com" {
		t.Fatalf("BatchRoot = %q", got)
	}
	if got := graph.BatchRoot(nil); got != "/me" {
		t.Fatalf("BatchRoot(nil) = %q", got)
	}
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{JSON: map[string]any{}} })
	if _, err := c.Do(context.Background(), "GET", graph.Root(ptr("a+b@example.com"))+"/messages", nil, nil); err != nil {
		t.Fatal(err)
	}
	if got := f.Last().RawPath; got != "/users/a%2Bb%40example.com/messages" {
		t.Fatalf("wire path = %q", got)
	}
	if got := f.Last().Path; got != "/users/a+b@example.com/messages" {
		t.Fatalf("decoded path = %q", got)
	}
}

func TestMeReachesTheWire(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{JSON: map[string]any{"id": "u1"}}
	})
	got, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if got["id"] != "u1" || f.Last().Path != "/me" {
		t.Fatalf("got %v at %s", got, f.Last().Path)
	}
}

func TestBearerHeader(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{JSON: map[string]any{}} })
	if _, err := c.Do(context.Background(), "GET", "/me", nil, nil); err != nil {
		t.Fatal(err)
	}
	if got := f.Last().Header.Get("Authorization"); got != "Bearer test-token" {
		t.Fatalf("Authorization = %q", got)
	}

	hint := errors.New("Not authenticated. Run `msgraph-mcp login` in a terminal to sign in.")
	c.Tokens = failingTokens{hint}
	_, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	if !errors.Is(err, hint) {
		t.Fatalf("err = %v, want the token source's error unchanged", err)
	}
}

type failingTokens struct{ err error }

func (f failingTokens) Token(context.Context) (string, error) { return "", f.err }

func TestQueryEncodesSpacesAsPercent20(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{JSON: map[string]any{}} })
	q := map[string][]string{"$search": {`"budget review"`}, "$top": {"25"}}
	if _, err := c.Do(context.Background(), "GET", "/me/messages", q, nil); err != nil {
		t.Fatal(err)
	}
	if got := f.Last().Query.Get("$search"); got != `"budget review"` {
		t.Fatalf("$search = %q", got)
	}
}

func TestBodyIsJSON(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{Status: 202} })
	got, err := c.Do(context.Background(), "POST", "/me/events/e1/accept", nil, map[string]any{"sendResponse": true})
	if err != nil || got != nil {
		t.Fatalf("got %v, %v; want nil, nil for 202", got, err)
	}
	r := f.Last()
	if r.Header.Get("Content-Type") != "application/json" || r.Body["sendResponse"] != true {
		t.Fatalf("request %+v", r)
	}
}

func TestDoGraphError(t *testing.T) {
	c, _ := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Error(404, "ErrorItemNotFound", "x")
	})
	_, err := c.Do(context.Background(), "GET", "/me/messages/m1", nil, nil)
	if err == nil || err.Error() != "Graph API 404: ErrorItemNotFound — x" {
		t.Fatalf("err = %v", err)
	}
	var api *graph.APIError
	if !errors.As(err, &api) || api.Status != 404 || api.Code != "ErrorItemNotFound" || api.Message != "x" {
		t.Fatalf("err = %#v", err)
	}
}

func TestDoErrorWithoutInner(t *testing.T) {
	c, _ := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Status: 500, JSON: map[string]any{"error": nil}}
	})
	_, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	if err == nil || err.Error() != "Graph API 500: Unknown — HTTP 500" {
		t.Fatalf("err = %v", err)
	}
}

func TestDoNonJSONError(t *testing.T) {
	c, _ := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Status: 502, Bytes: []byte("<html>Bad Gateway</html>"), Header: map[string]string{"Content-Type": "text/html"}}
	})
	_, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	if err == nil || err.Error() != "Graph API 502: Unknown — HTTP 502" {
		t.Fatalf("err = %v", err)
	}
}

func TestDoTransportError(t *testing.T) {
	c := graph.New(staticTokens("t"))
	c.BaseURL = "http://127.0.0.1:1/v1.0"
	c.MaxAttempts = 1
	_, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	var api *graph.APIError
	if !errors.As(err, &api) || api.Status != 0 || api.Code != "Unknown" || api.Message == "" {
		t.Fatalf("err = %#v", err)
	}
}

type staticTokens string

func (s staticTokens) Token(context.Context) (string, error) { return string(s), nil }

func TestDoRetriesThrottled(t *testing.T) {
	calls := 0
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		calls++
		if calls == 1 {
			return graphtest.Response{Status: 429, Header: map[string]string{"Retry-After": "2"}, JSON: map[string]any{"error": map[string]any{"code": "TooManyRequests", "message": "slow down"}}}
		}
		return graphtest.Response{JSON: map[string]any{"ok": true}}
	})
	got, err := c.Do(context.Background(), "GET", "/me/chats", nil, nil)
	if err != nil || got["ok"] != true {
		t.Fatalf("got %v, %v", got, err)
	}
	if len(f.Sleeps) != 1 || f.Sleeps[0] != 2*time.Second {
		t.Fatalf("sleeps = %v, want [2s]", f.Sleeps)
	}
}

func TestDoBacksOffWithoutRetryAfter(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Status: http.StatusServiceUnavailable, Header: map[string]string{"Retry-After": "soon"}}
	})
	_, err := c.Do(context.Background(), "GET", "/me", nil, nil)
	if err == nil || err.Error() != "Graph API 503: Unknown — HTTP 503" {
		t.Fatalf("err = %v", err)
	}
	want := []time.Duration{time.Second, 2 * time.Second, 4 * time.Second}
	if len(f.Sleeps) != len(want) {
		t.Fatalf("sleeps = %v, want %v", f.Sleeps, want)
	}
	for i := range want {
		if f.Sleeps[i] != want[i] {
			t.Fatalf("sleeps = %v, want %v", f.Sleeps, want)
		}
	}
	if len(f.Requests) != 4 {
		t.Fatalf("%d attempts, want 4", len(f.Requests))
	}
}

func TestNoRetryOnOtherErrors(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Error(400, "BadRequest", "no") })
	_, _ = c.Do(context.Background(), "GET", "/me", nil, nil)
	if len(f.Requests) != 1 || len(f.Sleeps) != 0 {
		t.Fatalf("requests %d sleeps %v", len(f.Requests), f.Sleeps)
	}
}

func TestAbsoluteNextLinkUsedAsIs(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response { return graphtest.Response{JSON: map[string]any{}} })
	next := "https://graph.microsoft.com/v1.0/me/messages?%24skip=25&%24top=25"
	if _, err := c.Do(context.Background(), "GET", next, nil, nil); err != nil {
		t.Fatal(err)
	}
	r := f.Last()
	if r.Path != "/me/messages" || r.Query.Get("$skip") != "25" {
		t.Fatalf("got %s ? %v", r.Path, r.Query)
	}
}

func TestGetBytes(t *testing.T) {
	c, f := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Response{Bytes: []byte{0x89, 'P', 'N', 'G'}, Header: map[string]string{"Content-Type": "application/octet-stream"}}
	})
	b, err := c.GetBytes(context.Background(), "/chats/c1/messages/m1/hostedContents/h1/$value")
	if err != nil || string(b) != "\x89PNG" {
		t.Fatalf("got %q, %v", b, err)
	}
	if f.Last().Path != "/chats/c1/messages/m1/hostedContents/h1/$value" {
		t.Fatalf("path %s", f.Last().Path)
	}
}
