package tools

import (
	"context"
	"encoding/json"
	"net/url"
	"strconv"
	"strings"

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

// strconvAtoi reads a JSON number (json.Number or float64) as an int.
func strconvAtoi(v any) (int, bool) {
	switch x := v.(type) {
	case json.Number:
		n, err := x.Int64()
		return int(n), err == nil
	case float64:
		return int(x), true
	case int:
		return x, true
	}
	return 0, false
}

// pyRepr quotes s the way Python's repr() does for ordinary strings.
func pyRepr(s string) string {
	if strings.Contains(s, "'") && !strings.Contains(s, `"`) {
		return `"` + s + `"`
	}
	return "'" + strings.ReplaceAll(s, "'", `\'`) + "'"
}
