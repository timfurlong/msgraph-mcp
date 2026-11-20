// Package config loads the server's settings from the process environment,
// a .env in the working directory, and ~/.msgraph-mcp/.env, in that order.
package config

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
)

// Scopes is every delegated permission the server asks for, all at once.
// ChannelMessage.Read.All needs tenant admin consent.
var Scopes = []string{
	"Mail.ReadWrite",
	"Mail.ReadWrite.Shared",
	"Mail.Send",
	"Calendars.ReadWrite",
	"Calendars.ReadWrite.Shared",
	"MailboxSettings.ReadWrite",
	"User.Read",
	// Teams read history
	"Chat.Read",
	"Team.ReadBasic.All",
	"Channel.ReadBasic.All",
	"ChannelMessage.Read.All",
}

// Config is the resolved configuration.
type Config struct {
	ClientID  string
	TenantID  string
	CachePath string
}

// Authority is the Entra sign-in authority for the tenant.
func (c Config) Authority() string {
	return "https://login.microsoftonline.com/" + c.TenantID
}

// Load reads the configuration for the running process.
func Load() (Config, error) {
	cwd, _ := os.Getwd()
	home, _ := os.UserHomeDir()
	return load(os.Getenv, cwd, home)
}

const legacyPrefix = "OUTLOOK_MCP_"

func load(getenv func(string) string, cwd, home string) (Config, error) {
	sources := []func(string) string{getenv}
	if cwd != "" {
		cwdEnv := readDotenv(filepath.Join(cwd, ".env"))
		sources = append(sources, func(k string) string { return cwdEnv[k] })
	}
	if home != "" {
		homeEnv := readDotenv(filepath.Join(home, ".msgraph-mcp", ".env"))
		sources = append(sources, func(k string) string { return homeEnv[k] })
	}
	get := func(k string) string {
		for _, src := range sources {
			if v := src(k); v != "" {
				return v
			}
		}
		return ""
	}
	// The new name wins over the legacy one wherever either is set.
	value := func(name string) string {
		if v := get(name); v != "" {
			return v
		}
		return get(strings.Replace(name, "MSGRAPH_MCP_", legacyPrefix, 1))
	}

	var c Config
	if c.ClientID = value("MSGRAPH_MCP_CLIENT_ID"); c.ClientID == "" {
		return Config{}, missing("MSGRAPH_MCP_CLIENT_ID")
	}
	if c.TenantID = value("MSGRAPH_MCP_TENANT_ID"); c.TenantID == "" {
		return Config{}, missing("MSGRAPH_MCP_TENANT_ID")
	}
	c.CachePath = cachePath(value("MSGRAPH_MCP_TOKEN_CACHE_PATH"), home)
	return c, nil
}

func missing(name string) error {
	return errors.New("Missing required env var: " + name +
		". Set it in the process env, a .env file in the working directory, or ~/.msgraph-mcp/.env.")
}

// cachePath honors an override (with ~ expanded), else the default location,
// else the pre-rename outlook-mcp cache when only that one exists.
func cachePath(override, home string) string {
	if override != "" {
		if override == "~" || strings.HasPrefix(override, "~/") {
			return filepath.Join(home, strings.TrimPrefix(override, "~"))
		}
		return override
	}
	def := filepath.Join(home, ".msgraph-mcp", "token_cache.bin")
	legacy := filepath.Join(home, ".outlook-mcp", "token_cache.bin")
	if _, err := os.Stat(def); err != nil {
		if _, err := os.Stat(legacy); err == nil {
			return legacy
		}
	}
	return def
}
