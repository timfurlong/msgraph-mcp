// Command msgraph-mcp is an MCP server for Microsoft Graph: Outlook mail and
// calendar, plus read-only Microsoft Teams.
package main

import (
	"fmt"
	"os"
)

// version is set at build time with -ldflags "-X main.version=...".
var version = "dev"

func main() {
	if len(os.Args) > 1 && (os.Args[1] == "--version" || os.Args[1] == "-v") {
		fmt.Println("msgraph-mcp", version)
		return
	}
	fmt.Fprintln(os.Stderr, "msgraph-mcp: server not built yet")
	os.Exit(1)
}
