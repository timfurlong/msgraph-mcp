.PHONY: build test install live

VERSION ?= $(shell git describe --tags --always --dirty 2>/dev/null || echo dev)
LDFLAGS := -s -w -X main.version=$(VERSION)

build:
	go build -ldflags='$(LDFLAGS)' -o msgraph-mcp .

test:
	go vet ./...
	go test ./...

install:
	mkdir -p $(HOME)/.local/bin
	go build -trimpath -ldflags='$(LDFLAGS)' -o $(HOME)/.local/bin/msgraph-mcp .

# Read-only calls plus a draft that is deleted, against the signed-in mailbox.
live:
	MSGRAPH_MCP_INTEGRATION=1 go test -count=1 -v ./internal/live/...
