package tools

import (
	"context"
	"net/url"
	"strconv"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

func encodeToken(next string) *string { return graph.EncodePageToken(next) }

// seg escapes one path segment (an id or a folder name).
func seg(s string) string { return graph.Segment(s) }

// fetchPage gets the first page at path with query q, or, when pageToken is
// set, the decoded nextLink as-is (its own query wins).
func fetchPage(ctx context.Context, g *graph.Client, path string, q url.Values, pageToken *string) (M, error) {
	if pageToken != nil {
		next, err := graph.DecodePageToken(*pageToken)
		if err != nil {
			return nil, err
		}
		return g.Do(ctx, "GET", next, nil, nil)
	}
	return g.Do(ctx, "GET", path, q, nil)
}

func itoa(n int) string { return strconv.Itoa(n) }
