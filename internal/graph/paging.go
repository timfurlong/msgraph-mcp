package graph

import (
	"encoding/base64"
	"net/url"
	"strings"
)

// EncodePageToken turns Graph's @odata.nextLink into an opaque token: the
// URL in URL-safe base64 without padding. An empty link gives nil.
func EncodePageToken(nextLink string) *string {
	if nextLink == "" {
		return nil
	}
	tok := base64.RawURLEncoding.EncodeToString([]byte(nextLink))
	return &tok
}

// DecodePageToken reverses EncodePageToken and refuses any URL that isn't
// https://graph.microsoft.com, so a token can't send the bearer token
// elsewhere.
func DecodePageToken(tok string) (string, error) {
	raw, err := base64.RawURLEncoding.DecodeString(strings.TrimRight(tok, "="))
	if err != nil {
		return "", Invalid("Invalid page_token: %v", err)
	}
	u := string(raw)
	parsed, err := url.Parse(u)
	if err != nil {
		return "", Invalid("Invalid page_token: %v", err)
	}
	if parsed.Scheme != "https" || parsed.Hostname() != "graph.microsoft.com" {
		got := "None"
		if h := parsed.Hostname(); h != "" {
			got = "'" + h + "'"
		}
		return "", Invalid("Invalid page_token: must point at graph.microsoft.com (got %s)", got)
	}
	return u, nil
}

// ValidateLimit checks 1 <= limit <= max.
func ValidateLimit(limit, max int) error {
	if limit < 1 || limit > max {
		return Invalid("limit must be an integer in 1..%d (got %d)", max, limit)
	}
	return nil
}
