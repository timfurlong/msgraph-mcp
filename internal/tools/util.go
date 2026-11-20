package tools

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

type whoamiIn struct {
	IncludeRaw bool `json:"include_raw"`
}

func registerUtil(s *mcp.Server, g *graph.Client) {
	add(s, "whoami", func(ctx context.Context, in whoamiIn) (any, error) {
		user, err := g.Do(ctx, "GET", "/me", nil, nil)
		if err != nil {
			return nil, err
		}
		return trim.User(user, in.IncludeRaw), nil
	})
}
