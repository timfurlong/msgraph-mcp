package auth

import (
	"context"
	"fmt"
	"io"

	"github.com/timfurlong/msgraph-mcp/internal/config"
)

// Login runs the device code flow: it prints the sign-in URL and code, waits
// for the user to finish in a browser, and saves the tokens to the cache.
func Login(ctx context.Context, cfg config.Config, out io.Writer) error {
	src, err := NewSource(cfg)
	if err != nil {
		return err
	}
	dc, err := src.app.AcquireTokenByDeviceCode(ctx, config.Scopes)
	if err != nil {
		return fmt.Errorf("failed to start device flow: %w", err)
	}
	fmt.Fprintln(out, dc.Result.Message)
	res, err := dc.AuthenticationResult(ctx)
	if err != nil {
		return fmt.Errorf("login failed: %w", err)
	}
	user := res.Account.PreferredUsername
	if user == "" {
		user = "<unknown>"
	}
	fmt.Fprintf(out, "Logged in as %s\n", user)
	fmt.Fprintf(out, "Token cache saved to %s\n", cfg.CachePath)
	return nil
}
