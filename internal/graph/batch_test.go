package graph_test

import (
	"context"
	"fmt"
	"testing"
	"time"

	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/graphtest"
)

// batchFake answers each $batch POST with script(callIndex, requests).
func batchFake(t *testing.T, script func(call int, reqs []any) []any) (*graph.Client, *graphtest.Fake) {
	call := 0
	return graphtest.New(t, func(r graphtest.Request) graphtest.Response {
		if r.Method != "POST" || r.Path != "/$batch" {
			t.Fatalf("unexpected %s %s", r.Method, r.Path)
		}
		reqs, _ := r.Body["requests"].([]any)
		resp := script(call, reqs)
		call++
		return graphtest.Response{JSON: map[string]any{"responses": resp}}
	})
}

func ok(id any) map[string]any {
	return map[string]any{"id": id, "status": 200, "body": map[string]any{"id": id}}
}

func echoOK(_ int, reqs []any) []any {
	var out []any
	for _, r := range reqs {
		out = append(out, ok(r.(map[string]any)["id"]))
	}
	return out
}

func reqs(n int) []graph.BatchRequest {
	var out []graph.BatchRequest
	for i := 0; i < n; i++ {
		out = append(out, graph.BatchRequest{ID: fmt.Sprintf("m%d", i), Method: "PATCH", URL: fmt.Sprintf("/me/messages/m%d", i), Body: map[string]any{"isRead": true}})
	}
	return out
}

func TestBatchEmptyInput(t *testing.T) {
	c, f := batchFake(t, echoOK)
	got, err := c.ExecuteBatch(context.Background(), nil)
	if err != nil || len(got) != 0 || len(f.Requests) != 0 {
		t.Fatalf("got %v, %v, %d posts", got, err, len(f.Requests))
	}
}

func TestBatchSingleChunkPreservesOrder(t *testing.T) {
	c, f := batchFake(t, func(_ int, reqs []any) []any {
		// Answer in reverse order.
		var out []any
		for i := len(reqs) - 1; i >= 0; i-- {
			out = append(out, ok(reqs[i].(map[string]any)["id"]))
		}
		return out
	})
	got, err := c.ExecuteBatch(context.Background(), reqs(3))
	if err != nil {
		t.Fatal(err)
	}
	for i, r := range got {
		if r.ID != fmt.Sprintf("m%d", i) || !r.OK || r.Status != 200 {
			t.Fatalf("result %d = %+v", i, r)
		}
	}
	sent := f.Last().Body["requests"].([]any)
	first := sent[0].(map[string]any)
	if first["id"] != "1" || first["method"] != "PATCH" || first["url"] != "/me/messages/m0" {
		t.Fatalf("sub-request %v", first)
	}
	if h := first["headers"].(map[string]any); h["Content-Type"] != "application/json" {
		t.Fatalf("headers %v", h)
	}
}

func TestBatchNoBodyNoContentType(t *testing.T) {
	c, f := batchFake(t, echoOK)
	_, err := c.ExecuteBatch(context.Background(), []graph.BatchRequest{{ID: "a", Method: "DELETE", URL: "/me/messages/a"}})
	if err != nil {
		t.Fatal(err)
	}
	sub := f.Last().Body["requests"].([]any)[0].(map[string]any)
	if _, has := sub["headers"]; has {
		t.Fatalf("headers sent without a body: %v", sub)
	}
	if _, has := sub["body"]; has {
		t.Fatalf("body sent: %v", sub)
	}
}

func TestBatchCallerHeadersMerge(t *testing.T) {
	c, f := batchFake(t, echoOK)
	_, _ = c.ExecuteBatch(context.Background(), []graph.BatchRequest{{ID: "a", Method: "POST", URL: "/x", Body: map[string]any{}, Headers: map[string]string{"Prefer": "y"}}})
	h := f.Last().Body["requests"].([]any)[0].(map[string]any)["headers"].(map[string]any)
	if h["Prefer"] != "y" || h["Content-Type"] != "application/json" {
		t.Fatalf("headers %v", h)
	}
}

func TestBatchChunksAt20(t *testing.T) {
	c, f := batchFake(t, echoOK)
	got, err := c.ExecuteBatch(context.Background(), reqs(45))
	if err != nil || len(got) != 45 {
		t.Fatalf("got %d, %v", len(got), err)
	}
	sizes := []int{}
	for _, r := range f.Requests {
		sizes = append(sizes, len(r.Body["requests"].([]any)))
	}
	if fmt.Sprint(sizes) != "[20 20 5]" {
		t.Fatalf("chunk sizes %v", sizes)
	}
	// Each chunk numbers from 1.
	if id := f.Requests[1].Body["requests"].([]any)[0].(map[string]any)["id"]; id != "1" {
		t.Fatalf("second chunk starts at %v", id)
	}
}

func TestBatchMixedSuccessFailure(t *testing.T) {
	c, _ := batchFake(t, func(int, []any) []any {
		return []any{
			ok("1"),
			map[string]any{"id": "2", "status": 404, "body": map[string]any{"error": map[string]any{"code": "ErrorItemNotFound", "message": "gone"}}},
			map[string]any{"id": "3", "status": 403, "body": map[string]any{"error": map[string]any{"code": "ErrorAccessDenied"}}},
			map[string]any{"id": "4", "status": 500, "body": "oops"},
		}
	})
	got, err := c.ExecuteBatch(context.Background(), reqs(4))
	if err != nil {
		t.Fatal(err)
	}
	if !got[0].OK || got[0].Body["id"] != "1" {
		t.Fatalf("0: %+v", got[0])
	}
	want := []string{"", "404 ErrorItemNotFound: gone", "403 ErrorAccessDenied", "500"}
	for i := 1; i < 4; i++ {
		if got[i].OK || got[i].Error != want[i] {
			t.Fatalf("%d: %+v, want error %q", i, got[i], want[i])
		}
	}
	if got[3].Body != nil {
		t.Fatalf("non-object body kept: %v", got[3].Body)
	}
}

func TestBatchMissingResponseMarkedFailed(t *testing.T) {
	c, _ := batchFake(t, func(int, []any) []any { return []any{ok("1")} })
	got, _ := c.ExecuteBatch(context.Background(), reqs(2))
	if got[1].OK || got[1].Status != 0 || got[1].Error != "No response from $batch" {
		t.Fatalf("%+v", got[1])
	}
}

func throttled(id string, retryAfter any) map[string]any {
	h := map[string]any{}
	if retryAfter != nil {
		h["Retry-After"] = retryAfter
	}
	return map[string]any{"id": id, "status": 429, "headers": h, "body": map[string]any{"error": map[string]any{"code": "TooManyRequests", "message": "slow"}}}
}

func TestBatch429RetriesOnlyThrottled(t *testing.T) {
	var retried []any
	c, f := batchFake(t, func(call int, reqs []any) []any {
		if call == 0 {
			return []any{ok("1"), throttled("2", "3"), ok("3"), throttled("4", "5"), throttled("5", "soon")}
		}
		retried = reqs
		return echoOK(call, reqs)
	})
	got, err := c.ExecuteBatch(context.Background(), reqs(5))
	if err != nil {
		t.Fatal(err)
	}
	for i, r := range got {
		if !r.OK || r.ID != fmt.Sprintf("m%d", i) {
			t.Fatalf("%d: %+v", i, r)
		}
	}
	if len(f.Sleeps) != 1 || f.Sleeps[0] != 5*time.Second {
		t.Fatalf("sleeps %v, want [5s]", f.Sleeps)
	}
	if len(retried) != 3 {
		t.Fatalf("retried %d, want 3", len(retried))
	}
	for i, want := range []string{"/me/messages/m1", "/me/messages/m3", "/me/messages/m4"} {
		sub := retried[i].(map[string]any)
		if sub["id"] != fmt.Sprint(i+1) || sub["url"] != want || sub["body"] == nil {
			t.Fatalf("retry %d = %v", i, sub)
		}
	}
}

func TestBatch429Persistent(t *testing.T) {
	c, f := batchFake(t, func(int, []any) []any { return []any{throttled("1", nil)} })
	got, _ := c.ExecuteBatch(context.Background(), reqs(1))
	if got[0].OK || got[0].Status != 429 || got[0].Error != "429 TooManyRequests: slow" {
		t.Fatalf("%+v", got[0])
	}
	if len(f.Requests) != 2 {
		t.Fatalf("%d posts, want 2 (one retry)", len(f.Requests))
	}
	if len(f.Sleeps) != 0 {
		t.Fatalf("slept %v with no Retry-After", f.Sleeps)
	}
}

func TestBatchWholeFailureIsGraphError(t *testing.T) {
	c, _ := graphtest.New(t, func(graphtest.Request) graphtest.Response {
		return graphtest.Error(401, "InvalidAuthenticationToken", "expired")
	})
	_, err := c.ExecuteBatch(context.Background(), reqs(1))
	if err == nil || err.Error() != "Graph API 401: InvalidAuthenticationToken — expired" {
		t.Fatalf("err = %v", err)
	}
}
