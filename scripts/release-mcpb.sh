#!/usr/bin/env bash
# Build the universal .mcpb bundle from GoReleaser's dist/ output.
#
#   scripts/release-mcpb.sh <version>
#
# An MCPB manifest picks a binary per OS but not per CPU, so the bundle holds
# the darwin universal binary plus linux and windows amd64. Prints the
# bundle's path and SHA-256.
set -euo pipefail

version="${1:?usage: release-mcpb.sh <version>}"
version="${version#v}"
root="$(cd "$(dirname "$0")/.." && pwd)"
dist="$root/dist"
stage="$(mktemp -d)"
trap 'rm -rf "$stage"' EXIT

find_bin() {
	# GoReleaser names build dirs <id>_<os>_<arch>[_<variant>]; the
	# universal binary's dir is <id>_darwin_all.
	local pattern="$1" name="$2" found
	found="$(find "$dist" -path "$dist/$pattern/$name" -type f | head -1)"
	if [ -z "$found" ]; then
		echo "release-mcpb: no $name under $dist/$pattern" >&2
		exit 1
	fi
	echo "$found"
}

mkdir -p "$stage/server"
cp "$(find_bin 'msgraph-mcp-universal_darwin_all' msgraph-mcp)" "$stage/server/msgraph-mcp-darwin"
cp "$(find_bin 'msgraph-mcp_linux_amd64*' msgraph-mcp)" "$stage/server/msgraph-mcp-linux"
cp "$(find_bin 'msgraph-mcp_windows_amd64*' msgraph-mcp.exe)" "$stage/server/msgraph-mcp.exe"
chmod +x "$stage/server/"*
cp "$root/LICENSE" "$stage/"
jq --arg v "$version" '.version = $v' "$root/mcpb/manifest.json" > "$stage/manifest.json"

out="$dist/msgraph-mcp.mcpb"
rm -f "$out"
npx -y @anthropic-ai/mcpb@2.1.2 validate "$stage/manifest.json" >&2
npx -y @anthropic-ai/mcpb@2.1.2 pack "$stage" "$out" >&2
sha="$(shasum -a 256 "$out" | awk '{print $1}')"
echo "$out $sha"
