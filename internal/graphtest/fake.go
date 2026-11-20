// Package graphtest is an HTTP-level fake of Microsoft Graph for tests. It
// records every request and answers from a scripted handler.
package graphtest

import (
	"bytes"
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

// Request is one recorded call. Path and RawPath have the /v1.0 prefix
// removed, so they read like Graph's documented paths ("/me/messages").
type Request struct {
	Method  string
	Path    string
	RawPath string
	Query   url.Values
	Body    map[string]any
	RawBody []byte
	Header  http.Header
}

// Response is what the handler returns. JSON is marshaled; Bytes is sent
// as-is; neither means an empty body. Status defaults to 200.
type Response struct {
	Status int
	JSON   any
	Bytes  []byte
	Header map[string]string
}

// Handler scripts the fake's answers.
type Handler func(r Request) Response

// Fake is a running fake Graph.
type Fake struct {
	mu       sync.Mutex
	Requests []Request
	Sleeps   []time.Duration
}

// Last returns the most recent request.
func (f *Fake) Last() Request {
	f.mu.Lock()
	defer f.mu.Unlock()
	return f.Requests[len(f.Requests)-1]
}

type staticToken string

func (s staticToken) Token(context.Context) (string, error) { return string(s), nil }

// rewrite sends every request for graph.microsoft.com to the fake server,
// so absolute nextLink URLs work too.
type rewrite struct{ target *url.URL }

func (r rewrite) RoundTrip(req *http.Request) (*http.Response, error) {
	if req.URL.Host == "graph.microsoft.com" {
		req = req.Clone(req.Context())
		req.URL.Scheme = r.target.Scheme
		req.URL.Host = r.target.Host
	}
	return http.DefaultTransport.RoundTrip(req)
}

// New starts a fake Graph and returns a client wired to it. Sleeps are
// recorded instead of slept.
func New(t *testing.T, h Handler) (*graph.Client, *Fake) {
	t.Helper()
	f := &Fake{}
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		body, _ := io.ReadAll(r.Body)
		req := Request{
			Method:  r.Method,
			Path:    strings.TrimPrefix(r.URL.Path, "/v1.0"),
			RawPath: strings.TrimPrefix(r.URL.EscapedPath(), "/v1.0"),
			Query:   r.URL.Query(),
			RawBody: body,
			Header:  r.Header.Clone(),
		}
		if len(body) > 0 {
			_ = json.Unmarshal(body, &req.Body)
		}
		f.mu.Lock()
		f.Requests = append(f.Requests, req)
		f.mu.Unlock()

		resp := h(req)
		for k, v := range resp.Header {
			w.Header().Set(k, v)
		}
		status := resp.Status
		if status == 0 {
			status = http.StatusOK
		}
		var out []byte
		switch {
		case resp.Bytes != nil:
			out = resp.Bytes
		case resp.JSON != nil:
			out, _ = json.Marshal(resp.JSON)
			if w.Header().Get("Content-Type") == "" {
				w.Header().Set("Content-Type", "application/json")
			}
		}
		w.WriteHeader(status)
		_, _ = io.Copy(w, bytes.NewReader(out))
	}))
	t.Cleanup(srv.Close)
	target, _ := url.Parse(srv.URL)

	c := graph.New(staticToken("test-token"))
	c.HTTP = &http.Client{Transport: rewrite{target: target}}
	c.Sleep = func(_ context.Context, d time.Duration) error {
		f.mu.Lock()
		f.Sleeps = append(f.Sleeps, d)
		f.mu.Unlock()
		return nil
	}
	return c, f
}

// Error is a Graph error body.
func Error(status int, code, message string) Response {
	return Response{Status: status, JSON: map[string]any{"error": map[string]any{"code": code, "message": message}}}
}
