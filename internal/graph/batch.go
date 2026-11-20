package graph

import (
	"context"
	"encoding/json"
	"strconv"
	"strings"
	"time"
)

// BatchRequest is one sub-request. ID is the caller's id; URL is relative to
// the API version ("/me/messages/x").
type BatchRequest struct {
	ID      string
	Method  string
	URL     string
	Body    any
	Headers map[string]string
}

// BatchResult is one sub-request's outcome, in input order.
type BatchResult struct {
	ID     string
	OK     bool
	Status int
	Error  string
	Body   map[string]any
}

const batchChunk = 20

// ExecuteBatch sends reqs through Graph's $batch endpoint in chunks of 20.
// Sub-requests that come back 429 are resent once, after sleeping the
// largest Retry-After among them. A failure of a whole $batch POST is
// returned as an error.
func (c *Client) ExecuteBatch(ctx context.Context, reqs []BatchRequest) ([]BatchResult, error) {
	results := map[string]BatchResult{}
	for start := 0; start < len(reqs); start += batchChunk {
		end := min(start+batchChunk, len(reqs))
		chunk := reqs[start:end]

		payload := make([]map[string]any, len(chunk))
		callerOf := map[string]string{}
		for i, r := range chunk {
			sub := strconv.Itoa(i + 1)
			callerOf[sub] = r.ID
			payload[i] = subRequest(sub, r)
		}
		responses, err := c.postBatch(ctx, payload)
		if err != nil {
			return nil, err
		}

		var throttled []map[string]any
		var wait time.Duration
		throttledIDs := map[string]bool{}
		for _, sub := range responses {
			if subStatus(sub) == 429 {
				throttledIDs[subID(sub)] = true
				if d, ok := subRetryAfter(sub); ok && d > wait {
					wait = d
				}
			}
		}
		for _, p := range payload {
			if throttledIDs[p["id"].(string)] {
				throttled = append(throttled, p)
			}
		}
		for _, sub := range responses {
			if throttledIDs[subID(sub)] {
				continue
			}
			if id, ok := callerOf[subID(sub)]; ok {
				results[id] = subResult(id, sub)
			}
		}
		if len(throttled) == 0 {
			continue
		}

		if wait > 0 {
			if err := c.Sleep(ctx, wait); err != nil {
				return nil, err
			}
		}
		retryCallerOf := map[string]string{}
		retry := make([]map[string]any, len(throttled))
		for j, orig := range throttled {
			sub := strconv.Itoa(j + 1)
			retryCallerOf[sub] = callerOf[orig["id"].(string)]
			cp := map[string]any{}
			for k, v := range orig {
				cp[k] = v
			}
			cp["id"] = sub
			retry[j] = cp
		}
		retryResponses, err := c.postBatch(ctx, retry)
		if err != nil {
			return nil, err
		}
		for _, sub := range retryResponses {
			if id, ok := retryCallerOf[subID(sub)]; ok {
				results[id] = subResult(id, sub)
			}
		}
	}

	out := make([]BatchResult, len(reqs))
	for i, r := range reqs {
		res, ok := results[r.ID]
		if !ok {
			res = BatchResult{ID: r.ID, Error: "No response from $batch"}
		}
		out[i] = res
	}
	return out, nil
}

func subRequest(id string, r BatchRequest) map[string]any {
	entry := map[string]any{"id": id, "method": r.Method, "url": r.URL}
	headers := map[string]any{}
	if r.Body != nil {
		entry["body"] = r.Body
		headers["Content-Type"] = "application/json"
	}
	for k, v := range r.Headers {
		headers[k] = v
	}
	if len(headers) > 0 {
		entry["headers"] = headers
	}
	return entry
}

func (c *Client) postBatch(ctx context.Context, payload []map[string]any) ([]map[string]any, error) {
	resp, err := c.Do(ctx, "POST", "/$batch", nil, map[string]any{"requests": payload})
	if err != nil {
		return nil, err
	}
	raw, _ := resp["responses"].([]any)
	out := make([]map[string]any, 0, len(raw))
	for _, r := range raw {
		if m, ok := r.(map[string]any); ok {
			out = append(out, m)
		}
	}
	return out, nil
}

func subID(sub map[string]any) string {
	switch v := sub["id"].(type) {
	case string:
		return v
	case json.Number:
		return v.String()
	}
	return ""
}

func subStatus(sub map[string]any) int {
	switch v := sub["status"].(type) {
	case json.Number:
		n, _ := v.Int64()
		return int(n)
	case float64:
		return int(v)
	}
	return 0
}

func subRetryAfter(sub map[string]any) (time.Duration, bool) {
	h, _ := sub["headers"].(map[string]any)
	var raw string
	switch v := h["Retry-After"].(type) {
	case string:
		raw = v
	case json.Number:
		raw = v.String()
	default:
		return 0, false
	}
	secs, err := strconv.ParseFloat(strings.TrimSpace(raw), 64)
	if err != nil || secs < 0 {
		return 0, false
	}
	d := time.Duration(secs * float64(time.Second))
	if d > maxRetryAfter {
		d = maxRetryAfter
	}
	return d, true
}

func subResult(id string, sub map[string]any) BatchResult {
	status := subStatus(sub)
	r := BatchResult{ID: id, Status: status, OK: status >= 200 && status < 300}
	body, isObj := sub["body"].(map[string]any)
	if isObj {
		r.Body = body
	}
	if !r.OK {
		r.Error = batchError(status, body, isObj)
	}
	return r
}

// batchError formats "status code: message", dropping missing parts.
func batchError(status int, body map[string]any, isObj bool) string {
	if !isObj {
		return strconv.Itoa(status)
	}
	e, _ := body["error"].(map[string]any)
	code, _ := e["code"].(string)
	msg, _ := e["message"].(string)
	prefix := strconv.Itoa(status)
	if code != "" {
		prefix += " " + code
	}
	if msg != "" {
		return prefix + ": " + msg
	}
	return prefix
}
