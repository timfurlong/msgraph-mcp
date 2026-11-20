package auth

import (
	"context"
	"encoding/json"
	"errors"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/cache"
	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/public"

	"github.com/timfurlong/msgraph-mcp/internal/config"
)

// A synthesized MSAL cache: three refresh tokens for one account and client,
// as MSAL Python leaves after scope changes, plus one for another client.
const threeTokens = `{
  "Account": {"acct": {"home_account_id": "uid.tid", "environment": "login.microsoftonline.com", "realm": "tid", "username": "user@example.com"}},
  "AccessToken": {},
  "IdToken": {},
  "AppMetadata": {"app": {"client_id": "client-a", "environment": "login.microsoftonline.com"}},
  "RefreshToken": {
    "rt-old":   {"home_account_id": "uid.tid", "environment": "login.microsoftonline.com", "client_id": "client-a", "credential_type": "RefreshToken", "secret": "s1", "target": "a", "last_modification_time": "1000"},
    "rt-new":   {"home_account_id": "uid.tid", "environment": "login.microsoftonline.com", "client_id": "client-a", "credential_type": "RefreshToken", "secret": "s3", "target": "a b", "last_modification_time": "3000"},
    "rt-mid":   {"home_account_id": "uid.tid", "environment": "login.microsoftonline.com", "client_id": "client-a", "credential_type": "RefreshToken", "secret": "s2", "target": "a b c", "last_modification_time": "2000"},
    "rt-other": {"home_account_id": "uid.tid", "environment": "login.microsoftonline.com", "client_id": "client-b", "credential_type": "RefreshToken", "secret": "s4", "last_modification_time": "500"}
  }
}`

func sections(t *testing.T, b []byte) map[string]map[string]map[string]any {
	t.Helper()
	var m map[string]map[string]map[string]any
	if err := json.Unmarshal(b, &m); err != nil {
		t.Fatal(err)
	}
	return m
}

func TestPruneKeepsNewestPerAccountClient(t *testing.T) {
	out, err := pruneRefreshTokens([]byte(threeTokens))
	if err != nil {
		t.Fatal(err)
	}
	got := sections(t, out)
	rts := got["RefreshToken"]
	if len(rts) != 2 || rts["rt-new"] == nil || rts["rt-other"] == nil {
		t.Fatalf("kept %v, want rt-new and rt-other", keys(rts))
	}
	in := sections(t, []byte(threeTokens))
	for _, s := range []string{"Account", "AccessToken", "IdToken", "AppMetadata"} {
		a, _ := json.Marshal(in[s])
		b, _ := json.Marshal(got[s])
		if string(a) != string(b) {
			t.Fatalf("section %s changed", s)
		}
	}
}

func TestStampAddsMissingTimestamps(t *testing.T) {
	in := `{"RefreshToken": {"go": {"client_id": "c", "home_account_id": "h", "environment": "e"}, "py": {"client_id": "c", "home_account_id": "h", "environment": "e", "last_modification_time": "1000"}}}`
	now := time.Unix(5000, 0)
	out, err := stampRefreshTokens([]byte(in), now)
	if err != nil {
		t.Fatal(err)
	}
	rts := sections(t, out)["RefreshToken"]
	if rts["go"]["last_modification_time"] != strconv.FormatInt(now.Unix(), 10) {
		t.Fatalf("go entry stamp = %v", rts["go"]["last_modification_time"])
	}
	if rts["py"]["last_modification_time"] != "1000" {
		t.Fatalf("py entry stamp changed to %v", rts["py"]["last_modification_time"])
	}
}

func TestGoWrittenTokenWinsAfterStamp(t *testing.T) {
	in := `{"RefreshToken": {"go": {"client_id": "c", "home_account_id": "h", "environment": "e"}, "py": {"client_id": "c", "home_account_id": "h", "environment": "e", "last_modification_time": "1000"}}}`
	stamped, err := stampRefreshTokens([]byte(in), time.Unix(5000, 0))
	if err != nil {
		t.Fatal(err)
	}
	out, err := pruneRefreshTokens(stamped)
	if err != nil {
		t.Fatal(err)
	}
	rts := sections(t, out)["RefreshToken"]
	if len(rts) != 1 || rts["go"] == nil {
		t.Fatalf("kept %v, want only go", keys(rts))
	}
}

type fakeMarshaler []byte

func (f fakeMarshaler) Marshal() ([]byte, error) { return f, nil }

type captureUnmarshaler struct{ got []byte }

func (c *captureUnmarshaler) Unmarshal(b []byte) error { c.got = b; return nil }

func TestExportWritesAtomicallyWithModes(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "cachedir")
	if err := os.Mkdir(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "token_cache.bin")
	fc := &FileCache{Path: path, Now: func() time.Time { return time.Unix(5000, 0) }}
	if err := fc.Export(context.Background(), fakeMarshaler(`{"RefreshToken":{}}`), cacheExportHints()); err != nil {
		t.Fatal(err)
	}
	fi, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if fi.Mode().Perm() != 0o600 {
		t.Fatalf("file mode %v, want 0600", fi.Mode().Perm())
	}
	di, _ := os.Stat(dir)
	if di.Mode().Perm() != 0o700 {
		t.Fatalf("dir mode %v, want 0700", di.Mode().Perm())
	}
	entries, _ := os.ReadDir(dir)
	if len(entries) != 1 {
		t.Fatalf("dir has %d entries, want only the cache file", len(entries))
	}
}

func TestReplaceMissingFileIsEmpty(t *testing.T) {
	fc := &FileCache{Path: filepath.Join(t.TempDir(), "nope", "token_cache.bin")}
	var u captureUnmarshaler
	if err := fc.Replace(context.Background(), &u, cacheReplaceHints()); err != nil {
		t.Fatalf("Replace: %v", err)
	}
	if string(u.got) != "{}" {
		t.Fatalf("unmarshaled %q from a missing file, want an empty cache", u.got)
	}
}

func TestReplacePrunesBeforeUnmarshal(t *testing.T) {
	path := filepath.Join(t.TempDir(), "token_cache.bin")
	if err := os.WriteFile(path, []byte(threeTokens), 0o600); err != nil {
		t.Fatal(err)
	}
	var u captureUnmarshaler
	if err := (&FileCache{Path: path}).Replace(context.Background(), &u, cacheReplaceHints()); err != nil {
		t.Fatal(err)
	}
	if n := len(sections(t, u.got)["RefreshToken"]); n != 2 {
		t.Fatalf("MSAL got %d refresh tokens, want 2", n)
	}
}

// noNetwork fails the test if MSAL makes any HTTP call.
type noNetwork struct{ t *testing.T }

func (n noNetwork) Do(r *http.Request) (*http.Response, error) {
	n.t.Errorf("unexpected network call: %s %s", r.Method, r.URL)
	return nil, errors.New("no network in tests")
}
func (n noNetwork) CloseIdleConnections() {}

func TestTokenNotAuthenticated(t *testing.T) {
	cases := map[string]*string{"missing file": nil, "empty file": ptr(""), "corrupt file": ptr("{not json")}
	for name, content := range cases {
		t.Run(name, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "token_cache.bin")
			if content != nil {
				if err := os.WriteFile(path, []byte(*content), 0o600); err != nil {
					t.Fatal(err)
				}
			}
			cfg := config.Config{ClientID: "00000000-0000-0000-0000-000000000000", TenantID: "common", CachePath: path}
			src, err := newSource(cfg, noNetwork{t})
			if err != nil {
				t.Fatal(err)
			}
			_, err = src.Token(context.Background())
			want := "Not authenticated. Run `msgraph-mcp login` in a terminal to sign in."
			if err == nil || err.Error() != want {
				t.Fatalf("err = %v, want %q", err, want)
			}
			if !errors.Is(err, ErrNotAuthenticated) {
				t.Fatal("error is not ErrNotAuthenticated")
			}
		})
	}
}

func ptr(s string) *string { return &s }

func keys[V any](m map[string]V) []string {
	var out []string
	for k := range m {
		out = append(out, k)
	}
	return out
}

func cacheExportHints() cache.ExportHints   { return cache.ExportHints{} }
func cacheReplaceHints() cache.ReplaceHints { return cache.ReplaceHints{} }

// oneAccount is a synthesized cache MSAL Go loads with no network: one
// account, its id token and a refresh token, for client 00000000-....
const oneAccount = `{
  "Account": {"uid.utid-login.microsoftonline.com-utid": {"home_account_id": "uid.utid", "environment": "login.microsoftonline.com", "realm": "utid", "local_account_id": "uid", "username": "user@example.com", "authority_type": "MSSTS"}},
  "RefreshToken": {"uid.utid-login.microsoftonline.com-refreshtoken-00000000-0000-0000-0000-000000000000--": {"home_account_id": "uid.utid", "environment": "login.microsoftonline.com", "credential_type": "RefreshToken", "client_id": "00000000-0000-0000-0000-000000000000", "secret": "s", "last_modification_time": "1000"}},
  "AccessToken": {}, "IdToken": {},
  "AppMetadata": {"appmetadata-login.microsoftonline.com-00000000-0000-0000-0000-000000000000": {"client_id": "00000000-0000-0000-0000-000000000000", "environment": "login.microsoftonline.com"}}
}`

func TestDeletedCacheSignsOutRunningServer(t *testing.T) {
	for name, gone := range map[string]func(string) error{
		"deleted": os.Remove,
		"emptied": func(p string) error { return os.WriteFile(p, nil, 0o600) },
	} {
		t.Run(name, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "token_cache.bin")
			if err := os.WriteFile(path, []byte(oneAccount), 0o600); err != nil {
				t.Fatal(err)
			}
			cfg := config.Config{ClientID: "00000000-0000-0000-0000-000000000000", TenantID: "utid", CachePath: path}
			src, err := newSource(cfg, noNetwork{t})
			if err != nil {
				t.Fatal(err)
			}
			accts, err := src.app.Accounts(context.Background())
			if err != nil || len(accts) != 1 {
				t.Fatalf("before: %d accounts, %v", len(accts), err)
			}
			if err := gone(path); err != nil {
				t.Fatal(err)
			}
			if _, err := src.Token(context.Background()); !errors.Is(err, ErrNotAuthenticated) {
				t.Fatalf("after %s: err = %v, want not authenticated", name, err)
			}
			if b, _ := os.ReadFile(path); len(b) > 0 {
				t.Fatalf("cache rewritten after %s: %s", name, b)
			}
		})
	}
}

const twoAccounts = `{
  "Account": {
    "a": {"home_account_id": "aaa.t", "environment": "login.microsoftonline.com", "realm": "t", "username": "a@example.com"},
    "b": {"home_account_id": "bbb.t", "environment": "login.microsoftonline.com", "realm": "t", "username": "b@example.com"}
  },
  "RefreshToken": {
    "rt-a": {"home_account_id": "aaa.t", "environment": "login.microsoftonline.com", "client_id": "c", "credential_type": "RefreshToken", "secret": "s", "last_modification_time": "1000"},
    "rt-b": {"home_account_id": "bbb.t", "environment": "login.microsoftonline.com", "client_id": "c", "credential_type": "RefreshToken", "secret": "s", "last_modification_time": "2000"}
  }
}`

func TestReplaceRecordsNewestAccount(t *testing.T) {
	path := filepath.Join(t.TempDir(), "token_cache.bin")
	if err := os.WriteFile(path, []byte(twoAccounts), 0o600); err != nil {
		t.Fatal(err)
	}
	fc := &FileCache{Path: path}
	var u captureUnmarshaler
	if err := fc.Replace(context.Background(), &u, cacheReplaceHints()); err != nil {
		t.Fatal(err)
	}
	if got := fc.preferredAccount(); got != "bbb.t" {
		t.Fatalf("preferred = %q, want bbb.t (newest refresh token)", got)
	}
}

func TestPickAccountIsDeterministic(t *testing.T) {
	accts := []public.Account{{HomeAccountID: "ccc.t"}, {HomeAccountID: "aaa.t"}, {HomeAccountID: "bbb.t"}}
	if got := pickAccount(accts, "bbb.t"); got.HomeAccountID != "bbb.t" {
		t.Fatalf("got %s, want the preferred account", got.HomeAccountID)
	}
	for i := 0; i < 5; i++ {
		if got := pickAccount([]public.Account{accts[i%3], accts[(i+1)%3], accts[(i+2)%3]}, ""); got.HomeAccountID != "aaa.t" {
			t.Fatalf("got %s, want aaa.t whatever the order", got.HomeAccountID)
		}
	}
}
