package graph_test

import (
	"context"
	"net/http"
	"net/http/httptest"
	"sync/atomic"
	"testing"
	"time"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

// dropping accepts each request, then closes the connection without a reply,
// as when a laptop sleeps after Graph has already acted on a POST.
func dropping(t *testing.T) (*graph.Client, *atomic.Int32) {
	var hits atomic.Int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		hits.Add(1)
		conn, _, err := w.(http.Hijacker).Hijack()
		if err == nil {
			conn.Close()
		}
	}))
	t.Cleanup(srv.Close)
	c := graph.New(staticTokens("t"))
	c.BaseURL = srv.URL + "/v1.0"
	c.Sleep = func(context.Context, time.Duration) error { return nil }
	return c, &hits
}

func TestTransportErrorNotRetriedForPOST(t *testing.T) {
	c, hits := dropping(t)
	_, err := c.Do(context.Background(), "POST", "/me/sendMail", nil, map[string]any{"message": map[string]any{}})
	if err == nil {
		t.Fatal("no error")
	}
	if n := hits.Load(); n != 1 {
		t.Fatalf("POST sent %d times, want 1", n)
	}
}

func TestTransportErrorRetriedForGET(t *testing.T) {
	c, hits := dropping(t)
	if _, err := c.Do(context.Background(), "GET", "/me", nil, nil); err == nil {
		t.Fatal("no error")
	}
	if n := hits.Load(); n != 4 {
		t.Fatalf("GET sent %d times, want 4", n)
	}
}
