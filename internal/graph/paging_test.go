package graph_test

import (
	"strings"
	"testing"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
)

func TestPageTokenRoundTrip(t *testing.T) {
	u := "https://graph.microsoft.com/v1.0/me/messages?$skip=25"
	tok := graph.EncodePageToken(u)
	if tok == nil || *tok == u || strings.Contains(*tok, "graph.microsoft.com") || strings.Contains(*tok, "=") {
		t.Fatalf("token %v", tok)
	}
	got, err := graph.DecodePageToken(*tok)
	if err != nil || got != u {
		t.Fatalf("decode = %q, %v", got, err)
	}
	if graph.EncodePageToken("") != nil {
		t.Fatal("empty nextLink should encode to nil")
	}
}

func TestPageTokenInvalid(t *testing.T) {
	_, err := graph.DecodePageToken("not-base64!")
	if err == nil || !strings.HasPrefix(err.Error(), "Invalid page_token: ") {
		t.Fatalf("err = %v", err)
	}
}

func TestPageTokenHostCheck(t *testing.T) {
	bad := graph.EncodePageToken("https://evil.example.com/foo")
	_, err := graph.DecodePageToken(*bad)
	if err == nil || err.Error() != "Invalid page_token: must point at graph.microsoft.com (got 'evil.example.com')" {
		t.Fatalf("err = %v", err)
	}
	insecure := graph.EncodePageToken("http://graph.microsoft.com/v1.0/me")
	if _, err := graph.DecodePageToken(*insecure); err == nil {
		t.Fatal("http scheme accepted")
	}
	nohost := graph.EncodePageToken("/v1.0/me")
	_, err = graph.DecodePageToken(*nohost)
	if err == nil || err.Error() != "Invalid page_token: must point at graph.microsoft.com (got None)" {
		t.Fatalf("err = %v", err)
	}
}

func TestValidateLimit(t *testing.T) {
	for _, ok := range []int{1, 25, 100} {
		if err := graph.ValidateLimit(ok, 100); err != nil {
			t.Fatalf("%d: %v", ok, err)
		}
	}
	err := graph.ValidateLimit(0, 100)
	if err == nil || err.Error() != "limit must be an integer in 1..100 (got 0)" {
		t.Fatalf("err = %v", err)
	}
	if graph.ValidateLimit(101, 100) == nil {
		t.Fatal("101 accepted")
	}
	if graph.ValidateLimit(50, 50) != nil {
		t.Fatal("50 rejected at max 50")
	}
	err = graph.ValidateLimit(51, 50)
	if err == nil || err.Error() != "limit must be an integer in 1..50 (got 51)" {
		t.Fatalf("err = %v", err)
	}
	var v *graph.ValidationError
	if !asValidation(err, &v) {
		t.Fatal("not a ValidationError")
	}
}
