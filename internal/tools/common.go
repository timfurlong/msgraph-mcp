package tools

import "github.com/timfurlong/msgraph-mcp/internal/graph"

func encodeToken(next string) *string { return graph.EncodePageToken(next) }
