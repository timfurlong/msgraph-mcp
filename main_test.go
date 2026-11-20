package main

import (
	"bytes"
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func buildBinary(t *testing.T) string {
	t.Helper()
	bin := filepath.Join(t.TempDir(), "msgraph-mcp")
	out, err := exec.Command("go", "build", "-o", bin, ".").CombinedOutput()
	if err != nil {
		t.Fatalf("build: %v\n%s", err, out)
	}
	return bin
}

func cleanEnv(home string, extra ...string) []string {
	env := []string{"HOME=" + home, "PATH=" + os.Getenv("PATH")}
	return append(env, extra...)
}

func TestBinaryListsTools(t *testing.T) {
	bin := buildBinary(t)
	home := t.TempDir()
	cmd := exec.Command(bin)
	cmd.Dir = home
	cmd.Env = cleanEnv(home, "MSGRAPH_MCP_CLIENT_ID=00000000-0000-0000-0000-000000000000", "MSGRAPH_MCP_TENANT_ID=common")
	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
	defer cancel()
	cs, err := mcp.NewClient(&mcp.Implementation{Name: "test"}, nil).Connect(ctx, &mcp.CommandTransport{Command: cmd}, nil)
	if err != nil {
		t.Fatal(err)
	}
	defer cs.Close()
	if !strings.Contains(cs.InitializeResult().Instructions, "Microsoft Graph") {
		t.Fatal("no instructions")
	}
	tools, err := cs.ListTools(ctx, nil)
	if err != nil || len(tools.Tools) != 50 {
		t.Fatalf("%d tools, %v", len(tools.Tools), err)
	}
	res, err := cs.CallTool(ctx, &mcp.CallToolParams{Name: "whoami", Arguments: map[string]any{}})
	if err != nil {
		t.Fatal(err)
	}
	text := res.Content[0].(*mcp.TextContent).Text
	if !res.IsError || text != "Not authenticated. Run `msgraph-mcp login` in a terminal to sign in." {
		t.Fatalf("whoami: %v %q", res.IsError, text)
	}
}

func TestBinaryFailsFastWithoutConfig(t *testing.T) {
	bin := buildBinary(t)
	home := t.TempDir()
	cmd := exec.Command(bin)
	cmd.Dir = home
	cmd.Env = cleanEnv(home)
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	err := cmd.Run()
	if err == nil {
		t.Fatal("exited 0 without config")
	}
	if !strings.Contains(stderr.String(), "msgraph-mcp: Missing required env var: MSGRAPH_MCP_CLIENT_ID") {
		t.Fatalf("stderr %q", stderr.String())
	}
}

func TestVersionFlag(t *testing.T) {
	out, err := exec.Command(buildBinary(t), "--version").Output()
	if err != nil || !strings.HasPrefix(string(out), "msgraph-mcp ") {
		t.Fatalf("%q %v", out, err)
	}
}
