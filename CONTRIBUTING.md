# Contributing

## Setup

```bash
git clone https://github.com/timfurlong/msgraph-mcp
cd msgraph-mcp
uv sync
cp .env.example .env   # fill in MSGRAPH_MCP_CLIENT_ID and MSGRAPH_MCP_TENANT_ID
uv run msgraph-mcp-login
```

Point an MCP host at the working tree with:

```bash
claude mcp add msgraph -- uv --directory "$(pwd)" run msgraph-mcp
```

## Checks

CI runs all four on every push and pull request, against Python 3.11, 3.12, and 3.13.

```bash
uv run pytest                            # unit tests
uv run ruff check .                      # lint
uv run ruff format .                     # format (CI runs --check)
uv run pyright src tests                 # types, must stay at zero errors
```

Integration tests make real Graph calls against the signed-in account and are skipped unless you opt in:

```bash
MSGRAPH_MCP_INTEGRATION=1 uv run pytest
```

## Releasing

Releases are tag-driven. Bump `version` in `pyproject.toml`, commit, then:

```bash
git tag vX.Y.Z && git push origin vX.Y.Z
```

The tag must match the `pyproject.toml` version or the workflow fails. From there
`.github/workflows/release.yml` runs the checks, publishes to PyPI via Trusted
Publishing (OIDC, no API tokens), and updates the MCP registry entry from
`server.json`.

Two things to know:

- PyPI versions can be yanked but never reused, so a bad release costs a version number.
- The README is the PyPI long description and the registry's ownership proof. It must
  keep the `mcp-name: io.github.timfurlong/msgraph-mcp` comment, and README changes only
  reach PyPI and the registry when a new version ships.
