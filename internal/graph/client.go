// Package graph is a small Microsoft Graph v1.0 client over net/http: bearer
// tokens, JSON in and out, Graph's error shape, retry on throttling, page
// tokens and $batch.
package graph

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"
)

// TokenSource hands out access tokens. Its errors reach callers unchanged.
type TokenSource interface {
	Token(ctx context.Context) (string, error)
}

// Client talks to Graph.
type Client struct {
	BaseURL     string
	HTTP        *http.Client
	Tokens      TokenSource
	Sleep       func(context.Context, time.Duration) error
	MaxAttempts int
}

const (
	maxBackoff    = 8 * time.Second
	maxRetryAfter = 60 * time.Second
)

// New returns a client for https://graph.microsoft.com/v1.0.
func New(ts TokenSource) *Client {
	return &Client{
		BaseURL:     "https://graph.microsoft.com/v1.0",
		HTTP:        &http.Client{Timeout: 2 * time.Minute},
		Tokens:      ts,
		Sleep:       sleep,
		MaxAttempts: 4,
	}
}

func sleep(ctx context.Context, d time.Duration) error {
	t := time.NewTimer(d)
	defer t.Stop()
	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-t.C:
		return nil
	}
}

// Root is the URL path for a mailbox: /me when mailbox is nil, else
// /users/{mailbox} with the mailbox escaped as a path segment.
func Root(mailbox *string) string {
	if mailbox == nil {
		return "/me"
	}
	return "/users/" + Segment(*mailbox)
}

// BatchRoot is Root for $batch sub-request URLs, which carry the mailbox
// unescaped.
func BatchRoot(mailbox *string) string {
	if mailbox == nil {
		return "/me"
	}
	return "/users/" + *mailbox
}

// Segment escapes s for use as one path segment, leaving only RFC 3986
// unreserved characters bare (as the Python SDK's URI templates did).
func Segment(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		c := s[i]
		if 'a' <= c && c <= 'z' || 'A' <= c && c <= 'Z' || '0' <= c && c <= '9' || c == '-' || c == '.' || c == '_' || c == '~' {
			b.WriteByte(c)
			continue
		}
		b.WriteString("%" + strings.ToUpper(strconv.FormatInt(int64(c)|0x100, 16)[1:]))
	}
	return b.String()
}

// Do sends a request and decodes a JSON object reply. path is relative to
// BaseURL, or an absolute https://graph.microsoft.com URL (a nextLink) used
// as-is. A 202/204 or empty reply returns nil, nil.
func (c *Client) Do(ctx context.Context, method, path string, query url.Values, body any) (map[string]any, error) {
	raw, err := c.send(ctx, method, path, query, body)
	if err != nil || len(bytes.TrimSpace(raw)) == 0 {
		return nil, err
	}
	dec := json.NewDecoder(bytes.NewReader(raw))
	dec.UseNumber()
	var out map[string]any
	if err := dec.Decode(&out); err != nil {
		return nil, &APIError{Status: 200, Code: "InvalidResponse", Message: err.Error()}
	}
	return out, nil
}

// GetBytes fetches raw bytes (a $value endpoint).
func (c *Client) GetBytes(ctx context.Context, path string) ([]byte, error) {
	return c.send(ctx, http.MethodGet, path, nil, nil)
}

func (c *Client) url(path string, query url.Values) string {
	u := path
	if !strings.HasPrefix(path, "https://") && !strings.HasPrefix(path, "http://") {
		u = c.BaseURL + path
	}
	if len(query) > 0 {
		sep := "?"
		if strings.Contains(u, "?") {
			sep = "&"
		}
		// Graph reads + in a query literally, so spaces go out as %20.
		u += sep + strings.ReplaceAll(query.Encode(), "+", "%20")
	}
	return u
}

func (c *Client) send(ctx context.Context, method, path string, query url.Values, body any) ([]byte, error) {
	var payload []byte
	if body != nil {
		var err error
		if payload, err = json.Marshal(body); err != nil {
			return nil, err
		}
	}
	target := c.url(path, query)
	attempts := c.MaxAttempts
	if attempts < 1 {
		attempts = 1
	}
	var lastErr error
	for attempt := 1; ; attempt++ {
		token, err := c.Tokens.Token(ctx)
		if err != nil {
			return nil, err
		}
		req, err := http.NewRequestWithContext(ctx, method, target, bytes.NewReader(payload))
		if err != nil {
			return nil, &APIError{Status: 0, Code: "Unknown", Message: err.Error()}
		}
		req.Header.Set("Authorization", "Bearer "+token)
		req.Header.Set("Accept", "application/json")
		if payload != nil {
			req.Header.Set("Content-Type", "application/json")
		}
		resp, err := c.HTTP.Do(req)
		if err != nil {
			lastErr = &APIError{Status: 0, Code: "Unknown", Message: err.Error()}
			if ctx.Err() != nil || attempt >= attempts {
				return nil, lastErr
			}
			if err := c.Sleep(ctx, backoff(attempt)); err != nil {
				return nil, lastErr
			}
			continue
		}
		data, readErr := io.ReadAll(resp.Body)
		resp.Body.Close()
		if resp.StatusCode >= 200 && resp.StatusCode < 300 {
			if readErr != nil {
				return nil, &APIError{Status: resp.StatusCode, Code: "Unknown", Message: readErr.Error()}
			}
			return data, nil
		}
		lastErr = apiError(resp.StatusCode, data)
		if !retryable(resp.StatusCode) || attempt >= attempts {
			return nil, lastErr
		}
		wait, ok := retryAfter(resp.Header.Get("Retry-After"))
		if !ok {
			wait = backoff(attempt)
		}
		if err := c.Sleep(ctx, wait); err != nil {
			return nil, lastErr
		}
	}
}

func retryable(status int) bool {
	return status == http.StatusTooManyRequests || status == http.StatusServiceUnavailable || status == http.StatusGatewayTimeout
}

// backoff is 1s, 2s, 4s, ... capped at maxBackoff.
func backoff(attempt int) time.Duration {
	d := time.Second << (attempt - 1)
	if d > maxBackoff || d <= 0 {
		return maxBackoff
	}
	return d
}

func retryAfter(h string) (time.Duration, bool) {
	secs, err := strconv.Atoi(strings.TrimSpace(h))
	if err != nil || secs < 0 {
		return 0, false
	}
	d := time.Duration(secs) * time.Second
	if d > maxRetryAfter {
		d = maxRetryAfter
	}
	return d, true
}

// apiError reads Graph's {"error": {"code", "message"}} body, falling back
// to code Unknown and "HTTP <status>".
func apiError(status int, body []byte) *APIError {
	e := &APIError{Status: status, Code: "Unknown", Message: "HTTP " + strconv.Itoa(status)}
	var parsed struct {
		Error *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if json.Unmarshal(body, &parsed) == nil && parsed.Error != nil {
		if parsed.Error.Code != "" {
			e.Code = parsed.Error.Code
		}
		if parsed.Error.Message != "" {
			e.Message = parsed.Error.Message
		}
	}
	return e
}

// IsStatus reports whether err is a Graph error with the given status.
func IsStatus(err error, status int) bool {
	var api *APIError
	return errors.As(err, &api) && api.Status == status
}
