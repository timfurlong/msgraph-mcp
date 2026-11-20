// Package auth signs in with MSAL Go (device code flow) and hands the server
// access tokens silently from the shared token cache.
package auth

import (
	"context"
	"errors"
	"net"
	"net/http"
	"net/url"
	"sort"
	"sync"

	msalerrors "github.com/AzureAD/microsoft-authentication-library-for-go/apps/errors"
	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/public"

	"github.com/timfurlong/msgraph-mcp/internal/config"
)

// LoginHint is what every tool returns when there is no usable sign-in.
const LoginHint = "Not authenticated. Run `msgraph-mcp login` in a terminal to sign in."

// NotAuthenticatedError means there is no account in the cache, or its
// refresh token no longer works.
type NotAuthenticatedError struct{}

func (NotAuthenticatedError) Error() string { return LoginHint }

// ErrNotAuthenticated is the sentinel for errors.Is.
var ErrNotAuthenticated error = NotAuthenticatedError{}

// Source hands out access tokens for the configured scopes.
type Source struct {
	app   public.Client
	cache *FileCache
	mu    sync.Mutex
}

// NewSource builds the MSAL public client over the token cache file.
func NewSource(cfg config.Config) (*Source, error) {
	return newSource(cfg, nil)
}

// httpDoer is MSAL's HTTP client interface; tests pass one that forbids
// network calls.
type httpDoer interface {
	Do(*http.Request) (*http.Response, error)
	CloseIdleConnections()
}

func newSource(cfg config.Config, hc httpDoer) (*Source, error) {
	fc := &FileCache{Path: cfg.CachePath}
	opts := []public.Option{
		public.WithAuthority(cfg.Authority()),
		public.WithCache(fc),
	}
	if hc != nil {
		opts = append(opts, public.WithHTTPClient(hc))
	}
	app, err := public.New(cfg.ClientID, opts...)
	if err != nil {
		return nil, err
	}
	return &Source{app: app, cache: fc}, nil
}

// Token returns an access token from the cache, refreshing it with the cached
// refresh token when needed. It never starts an interactive sign-in.
func (s *Source) Token(ctx context.Context) (string, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	accts, err := s.app.Accounts(ctx)
	if err != nil || len(accts) == 0 {
		return "", ErrNotAuthenticated
	}
	acct := pickAccount(accts, s.cache.preferredAccount())
	res, err := s.app.AcquireTokenSilent(ctx, config.Scopes, public.WithSilentAccount(acct))
	if err != nil {
		if isNetworkError(err) {
			return "", err
		}
		return "", ErrNotAuthenticated
	}
	return res.AccessToken, nil
}

// isNetworkError reports a failure to reach Entra at all, as opposed to Entra
// refusing the refresh token. Only the latter calls for a new sign-in.
func isNetworkError(err error) bool {
	var callErr msalerrors.CallErr
	if errors.As(err, &callErr) && callErr.Resp != nil {
		return false
	}
	var urlErr *url.Error
	var netErr net.Error
	return errors.As(err, &urlErr) || errors.As(err, &netErr)
}

// pickAccount chooses the account to act as when the cache holds more than
// one: the one that signed in or refreshed most recently, else the lowest
// home account id. MSAL Go returns accounts in map order, so taking the first
// would switch mailboxes between calls.
func pickAccount(accts []public.Account, preferred string) public.Account {
	sorted := append([]public.Account(nil), accts...)
	sort.Slice(sorted, func(i, j int) bool { return sorted[i].HomeAccountID < sorted[j].HomeAccountID })
	for _, a := range sorted {
		if a.HomeAccountID == preferred {
			return a
		}
	}
	return sorted[0]
}
