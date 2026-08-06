#!/usr/bin/env bash
# Build script for Office LLM Harness
# Builds both the .NET MCP server and the Office JS PowerPoint add-in.
#
# Usage:
#   ./scripts/build.sh          # Build both (production)
#   ./scripts/build.sh dev      # Build add-in for development
#   ./scripts/build.sh mcp      # Build MCP server only
#   ./scripts/build.sh addin    # Build PowerPoint add-in only

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

log() { echo -e "${GREEN}[build]${NC} $*"; }
warn() { echo -e "${YELLOW}[warn]${NC} $*"; }
err() {
	echo -e "${RED}[error]${NC} $*" >&2
	exit 1
}

# --- Build Office JS Add-in ---
build_addin() {
	local mode="${1:-production}"
	log "Building PowerPoint add-in ($mode)..."

	cd "$PROJECT_ROOT/src/office-addin"

	# Install dependencies if needed
	if [ ! -d "node_modules" ]; then
		log "Installing npm dependencies..."
		npm ci --prefer-offline 2>/dev/null || npm install
	fi

	if [ "$mode" = "production" ]; then
		npx webpack --mode production
		log "PowerPoint add-in built → dist/"
	else
		npx webpack --mode development
		log "PowerPoint add-in built (dev) → dist/"
	fi
}

# --- Build .NET MCP Server ---
build_mcp() {
	local runtime="${1:-}"

	log "Building MCP server..."

	cd "$PROJECT_ROOT/src/mcp-server"

	# Restore NuGet packages
	dotnet restore --verbosity quiet

	if [ -n "$runtime" ]; then
		# Self-contained for specific platform
		dotnet publish -c Release -r "$runtime" --self-contained true -p:PublishSingleFile=true -o "publish/$runtime"
		log "MCP server built → publish/$runtime/"
	else
		# Framework-dependent (needs .NET runtime on target)
		dotnet publish -c Release -p:PublishSingleFile=false -o "publish"
		log "MCP server built (framework-dependent) → publish/"
	fi
}

# --- Main ---
TARGET="${1:-all}"
MODE="${2:-production}"

case "$TARGET" in
all | addin)
	build_addin "$MODE"
	;;
mcp)
	# Detect target platform
	if [[ "$(uname)" == "Linux" ]]; then
		build_mcp "linux-x64"
	elif [[ "$(uname)" == "Darwin" ]]; then
		build_mcp "osx-arm64"
	else
		build_mcp "win-x64"
	fi
	;;
dev)
	build_addin "development"
	;;
*)
	err "Unknown target: $TARGET. Use: all, mcp, addin, dev"
	;;
esac

log "Build complete!"
