package config

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

// env returns a getenv over a fixed map.
func env(m map[string]string) func(string) string {
	return func(k string) string { return m[k] }
}

func writeFile(t *testing.T, path, content string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o700); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(content), 0o600); err != nil {
		t.Fatal(err)
	}
}

func dirs(t *testing.T) (cwd, home string) {
	t.Helper()
	return t.TempDir(), t.TempDir()
}

func TestLoadFromProcessEnv(t *testing.T) {
	cwd, home := dirs(t)
	c, err := load(env(map[string]string{"MSGRAPH_MCP_CLIENT_ID": "fake-client-id", "MSGRAPH_MCP_TENANT_ID": "abc-tenant"}), cwd, home)
	if err != nil {
		t.Fatal(err)
	}
	if c.ClientID != "fake-client-id" || c.TenantID != "abc-tenant" {
		t.Fatalf("got %+v", c)
	}
	if got := c.Authority(); got != "https://login.microsoftonline.com/abc-tenant" {
		t.Fatalf("authority = %q", got)
	}
}

func TestLegacyNames(t *testing.T) {
	cwd, home := dirs(t)
	c, err := load(env(map[string]string{"OUTLOOK_MCP_CLIENT_ID": "legacy-client-id", "OUTLOOK_MCP_TENANT_ID": "legacy-tenant", "OUTLOOK_MCP_TOKEN_CACHE_PATH": "/x/cache.bin"}), cwd, home)
	if err != nil {
		t.Fatal(err)
	}
	if c.ClientID != "legacy-client-id" || c.TenantID != "legacy-tenant" || c.CachePath != "/x/cache.bin" {
		t.Fatalf("got %+v", c)
	}
}

func TestLoadPrecedence(t *testing.T) {
	cwd, home := dirs(t)
	writeFile(t, filepath.Join(cwd, ".env"), "MSGRAPH_MCP_CLIENT_ID=from-cwd\nMSGRAPH_MCP_TENANT_ID=tenant-cwd\n")
	writeFile(t, filepath.Join(home, ".msgraph-mcp", ".env"), "MSGRAPH_MCP_CLIENT_ID=from-home\nMSGRAPH_MCP_TENANT_ID=tenant-home\nMSGRAPH_MCP_TOKEN_CACHE_PATH=/home/cache.bin\n")

	c, err := load(env(map[string]string{"MSGRAPH_MCP_CLIENT_ID": "from-process"}), cwd, home)
	if err != nil {
		t.Fatal(err)
	}
	if c.ClientID != "from-process" {
		t.Fatalf("client = %q, want process env to win", c.ClientID)
	}
	if c.TenantID != "tenant-cwd" {
		t.Fatalf("tenant = %q, want cwd .env over home .env", c.TenantID)
	}
	if c.CachePath != "/home/cache.bin" {
		t.Fatalf("cache = %q, want value from home .env", c.CachePath)
	}
}

func TestLoadUserDotenvOnly(t *testing.T) {
	cwd, home := dirs(t)
	writeFile(t, filepath.Join(home, ".msgraph-mcp", ".env"), "# ids\nexport MSGRAPH_MCP_CLIENT_ID=\"quoted-id\"\nMSGRAPH_MCP_TENANT_ID='single'\n\n")
	c, err := load(env(nil), cwd, home)
	if err != nil {
		t.Fatal(err)
	}
	if c.ClientID != "quoted-id" || c.TenantID != "single" {
		t.Fatalf("got %+v", c)
	}
}

func TestMissingMessage(t *testing.T) {
	cwd, home := dirs(t)
	_, err := load(env(map[string]string{"MSGRAPH_MCP_TENANT_ID": "t"}), cwd, home)
	want := "Missing required env var: MSGRAPH_MCP_CLIENT_ID. Set it in the process env, a .env file in the working directory, or ~/.msgraph-mcp/.env."
	if err == nil || err.Error() != want {
		t.Fatalf("err = %v, want %q", err, want)
	}
	_, err = load(env(map[string]string{"MSGRAPH_MCP_CLIENT_ID": "c"}), cwd, home)
	if err == nil || err.Error() != "Missing required env var: MSGRAPH_MCP_TENANT_ID. Set it in the process env, a .env file in the working directory, or ~/.msgraph-mcp/.env." {
		t.Fatalf("err = %v", err)
	}
}

func base() map[string]string {
	return map[string]string{"MSGRAPH_MCP_CLIENT_ID": "c", "MSGRAPH_MCP_TENANT_ID": "t"}
}

func TestCachePathDefault(t *testing.T) {
	cwd, home := dirs(t)
	c, err := load(env(base()), cwd, home)
	if err != nil {
		t.Fatal(err)
	}
	if want := filepath.Join(home, ".msgraph-mcp", "token_cache.bin"); c.CachePath != want {
		t.Fatalf("cache = %q, want %q", c.CachePath, want)
	}
}

func TestCachePathFallsBackToLegacy(t *testing.T) {
	cwd, home := dirs(t)
	legacy := filepath.Join(home, ".outlook-mcp", "token_cache.bin")
	writeFile(t, legacy, "cache")
	c, _ := load(env(base()), cwd, home)
	if c.CachePath != legacy {
		t.Fatalf("cache = %q, want legacy %q", c.CachePath, legacy)
	}
}

func TestCachePathPrefersNewOverLegacy(t *testing.T) {
	cwd, home := dirs(t)
	newPath := filepath.Join(home, ".msgraph-mcp", "token_cache.bin")
	writeFile(t, newPath, "cache")
	writeFile(t, filepath.Join(home, ".outlook-mcp", "token_cache.bin"), "cache")
	c, _ := load(env(base()), cwd, home)
	if c.CachePath != newPath {
		t.Fatalf("cache = %q, want %q", c.CachePath, newPath)
	}
}

func TestCachePathOverrideExpandsTilde(t *testing.T) {
	cwd, home := dirs(t)
	m := base()
	m["MSGRAPH_MCP_TOKEN_CACHE_PATH"] = "~/custom/cache.bin"
	c, _ := load(env(m), cwd, home)
	if want := filepath.Join(home, "custom", "cache.bin"); c.CachePath != want {
		t.Fatalf("cache = %q, want %q", c.CachePath, want)
	}
}

func TestScopes(t *testing.T) {
	want := []string{
		"Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.Send",
		"Calendars.ReadWrite", "Calendars.ReadWrite.Shared",
		"MailboxSettings.ReadWrite", "User.Read",
		"Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All",
	}
	if !reflect.DeepEqual(Scopes, want) {
		t.Fatalf("Scopes = %v", Scopes)
	}
}
