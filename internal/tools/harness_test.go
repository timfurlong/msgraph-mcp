package tools

import (
	"context"
	"encoding/json"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

// newSession starts a server with every tool over a fake Graph and returns a
// connected client session.
func newSession(t *testing.T, h graphtest.Handler) (*mcp.ClientSession, *graphtest.Fake) {
	t.Helper()
	g, f := graphtest.New(t, h)
	return sessionFor(t, g), f
}

func sessionFor(t *testing.T, g *graph.Client) *mcp.ClientSession {
	t.Helper()
	ctx := context.Background()
	srv := mcp.NewServer(&mcp.Implementation{Name: "msgraph-mcp", Version: "test"}, nil)
	Register(srv, g)
	ct, st := mcp.NewInMemoryTransports()
	if _, err := srv.Connect(ctx, st, nil); err != nil {
		t.Fatal(err)
	}
	cs, err := mcp.NewClient(&mcp.Implementation{Name: "test"}, nil).Connect(ctx, ct, nil)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { cs.Close() })
	return cs
}

// callTool calls a tool and returns the first text block, whether it is an
// error, and the whole result.
func callTool(t *testing.T, cs *mcp.ClientSession, name string, args map[string]any) (string, bool, *mcp.CallToolResult) {
	t.Helper()
	if args == nil {
		args = map[string]any{}
	}
	res, err := cs.CallTool(context.Background(), &mcp.CallToolParams{Name: name, Arguments: args})
	if err != nil {
		t.Fatalf("%s: protocol error: %v", name, err)
	}
	text := ""
	for _, c := range res.Content {
		if tc, ok := c.(*mcp.TextContent); ok {
			text = tc.Text
			break
		}
	}
	return text, res.IsError, res
}

// callJSON calls a tool that must succeed and decodes its JSON text.
func callJSON(t *testing.T, cs *mcp.ClientSession, name string, args map[string]any) map[string]any {
	t.Helper()
	text, isErr, _ := callTool(t, cs, name, args)
	if isErr {
		t.Fatalf("%s returned error: %s", name, text)
	}
	var out map[string]any
	if err := json.Unmarshal([]byte(text), &out); err != nil {
		t.Fatalf("%s: not a JSON object: %v\n%s", name, err, text)
	}
	return out
}

// callErr calls a tool that must fail and returns the error text.
func callErr(t *testing.T, cs *mcp.ClientSession, name string, args map[string]any) string {
	t.Helper()
	text, isErr, _ := callTool(t, cs, name, args)
	if !isErr {
		t.Fatalf("%s succeeded, want error; got %s", name, text)
	}
	return text
}

func jsonOK(v any) graphtest.Response { return graphtest.Response{JSON: v} }

func jsonUnmarshal(s string, v any) error { return json.Unmarshal([]byte(s), v) }
