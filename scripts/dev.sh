#!/usr/bin/env bash
# Development server for Office LLM Harness
# Starts the webpack dev server for the PowerPoint add-in.
#
# Usage:
#   ./scripts/dev.sh

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "Starting Office LLM Harness dev server..."
echo "Add-in will be available at: http://127.0.0.1:3000/taskpane.html"
echo ""
echo "To test in PowerPoint:"
echo "  1. Open PowerPoint on Windows"
echo "  2. File → Options → Trust Center → Trust Center Settings"
echo "  3. Trusted Add-in Publishers → Add the manifest location"
echo "  4. Or use 'sideload' via: https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-add-in-for-word"
echo ""
echo "Press Ctrl+C to stop."
echo "---"

cd "$PROJECT_ROOT/src/office-addin"
npx webpack serve --mode development --host 127.0.0.1 --port 3000
