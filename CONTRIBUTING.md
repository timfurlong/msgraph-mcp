# Contributing

## Setup

You need Go 1.26 or later (`go.mod`'s `toolchain` line fetches it automatically with `GOTOOLCHAIN=auto`, the default).

```bash
git clone https://github.com/timfurlong/msgraph-mcp
cd msgraph-mcp
cp .env.example .env   # fill in MSGRAPH_MCP_CLIENT_ID and MSGRAPH_MCP_TENANT_ID
go run . login
```

Point an MCP host at a build of the working tree with:

```bash
make install           # builds into ~/.local/bin/msgraph-mcp
claude mcp add msgraph -- ~/.local/bin/msgraph-mcp
```

## Layout

- `main.go`: the binary (`login`, `--version`, otherwise the stdio server).
- `internal/config`: environment and `.env` loading.
- `internal/auth`: MSAL Go sign-in and the shared token cache file.
- `internal/graph`: a small Graph client over `net/http` (errors, retry, page tokens, `$batch`), plus `internal/graphtest`, an HTTP-level fake of Graph for tests.
- `internal/trim`: reshapes Graph JSON into the trimmed objects the tools return.
- `internal/tools`: the 50 tools. Their names, descriptions and input schemas live in `internal/tools/toolspec/tools.json`; edit a tool's description or arguments there.

## Checks

CI runs these on every push and pull request:

```bash
make test              # go vet + go test
gofmt -l .             # must print nothing
```

The live suite makes real Graph calls as the signed-in account: read-only calls plus one draft it creates and deletes. It's skipped unless you opt in:

```bash
make live
```

## Releasing

Releases are tag-driven. Bump `version` in `server.json`, commit, then:

```bash
git tag vX.Y.Z && git push origin vX.Y.Z
```

The tag must match `server.json`'s version or the workflow fails. From there `.github/workflows/release.yml` runs the checks, builds the binaries with GoReleaser and publishes the GitHub release, attaches the `msgraph-mcp.mcpb` bundle, and publishes the MCP registry entry from `server.json` (GitHub OIDC, no tokens). A tag can't be reused once its release is public, so fix a bad release forward with a new patch version.
