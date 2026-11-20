package tools

import (
	"bytes"
	"encoding/json"
	"strings"

	"github.com/timfurlong/msgraph-mcp/internal/trim"
)

// M is a JSON object.
type M = trim.M

// jsonText renders v as JSON indented two spaces, without escaping <, > and
// & (message bodies are HTML).
func jsonText(v any) (string, error) {
	var buf bytes.Buffer
	enc := json.NewEncoder(&buf)
	enc.SetEscapeHTML(false)
	enc.SetIndent("", "  ")
	if err := enc.Encode(v); err != nil {
		return "", err
	}
	return strings.TrimSuffix(buf.String(), "\n"), nil
}

// page is the shape every paged tool returns.
func page(items []any, nextLink any) M {
	next, _ := nextLink.(string)
	var tok any
	if t := encodeToken(next); t != nil {
		tok = *t
	}
	return M{"items": items, "next_page_token": tok}
}

// items returns resp["value"] as objects.
func items(resp M) []M {
	raw, _ := resp["value"].([]any)
	out := make([]M, 0, len(raw))
	for _, v := range raw {
		if m, ok := v.(M); ok {
			out = append(out, m)
		}
	}
	return out
}
