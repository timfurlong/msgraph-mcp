// Command msgraph-mcp is an MCP server for Microsoft Graph: Outlook mail and
// calendar, plus read-only Microsoft Teams. Run with no arguments it serves
// MCP over stdio; "msgraph-mcp login" signs in with a device code.
package main

import (
	"context"
	"fmt"
	"os"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/timfurlong/msgraph-mcp/internal/auth"
	"github.com/timfurlong/msgraph-mcp/internal/config"
	"github.com/timfurlong/msgraph-mcp/internal/graph"
	"github.com/timfurlong/msgraph-mcp/internal/server"
)

// version is set at build time with -ldflags "-X main.version=...".
var version = "dev"

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, "msgraph-mcp:", err)
		os.Exit(1)
	}
}

func run(args []string) error {
	if len(args) > 0 {
		switch args[0] {
		case "--version", "-v", "version":
			fmt.Println("msgraph-mcp", version)
			return nil
		case "login":
			cfg, err := config.Load()
			if err != nil {
				return err
			}
			return auth.Login(context.Background(), cfg, os.Stdout)
		default:
			return fmt.Errorf("unknown command %q (run with no arguments to serve MCP over stdio, or \"login\" to sign in)", args[0])
		}
	}
	// Fail fast on missing configuration, before serving.
	cfg, err := config.Load()
	if err != nil {
		return err
	}
	tokens, err := auth.NewSource(cfg)
	if err != nil {
		return err
	}
	srv := server.New(graph.New(tokens), version)
	return srv.Run(context.Background(), &mcp.StdioTransport{})
}
