// Package tools registers the 50 Microsoft Graph tools. Each tool's name,
// description and input schema come from the toolspec snapshot, so clients
// see exactly what the Python server exposed.
package tools

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/tools/toolspec"
)

// Register adds every tool to s, in the Python server's order.
func Register(s *mcp.Server, g *graph.Client) {
	registerUtil(s, g)
	registerMailRead(s, g)
	registerMailWrite(s, g)
	registerMailFolders(s, g)
	registerMailActions(s, g)
	registerMailBatch(s, g)
	registerMailRules(s, g)
}

// add registers one tool. The handler returns a JSON-able value (sent as an
// indented JSON text block) or a ready *mcp.CallToolResult. A returned error
// becomes an isError result carrying the error's text.
func add[In any](s *mcp.Server, name string, h func(ctx context.Context, in In) (any, error)) {
	spec, ok := toolspec.Get(name)
	if !ok {
		panic("tools: " + name + " is not in the tool snapshot")
	}
	tool := &mcp.Tool{Name: spec.Name, Description: spec.Description, InputSchema: spec.InputSchema}
	mcp.AddTool(s, tool, func(ctx context.Context, _ *mcp.CallToolRequest, in In) (*mcp.CallToolResult, any, error) {
		out, err := h(ctx, in)
		if err != nil {
			return nil, nil, err
		}
		if res, ok := out.(*mcp.CallToolResult); ok {
			return res, nil, nil
		}
		text, err := jsonText(out)
		if err != nil {
			return nil, nil, err
		}
		return &mcp.CallToolResult{Content: []mcp.Content{&mcp.TextContent{Text: text}}}, nil, nil
	})
}
