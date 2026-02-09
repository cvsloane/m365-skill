#!/usr/bin/env bash
# =============================================================================
# List Available MS 365 Tools
# =============================================================================
# This script displays all available Microsoft 365 tools exposed through the
# MCP server. Useful for discovering capabilities and debugging.
#
# Usage:
#   ./list-tools.sh
#
# Requirements:
#   - mcporter must be installed and configured
#   - @softeria/ms-365-mcp-server must be installed
#
# See also: README.md for detailed tool descriptions
# =============================================================================

set -euo pipefail

echo "Available MS 365 Tools"
echo "======================"
echo ""

# Attempt to list tools with schema information
if mcporter list ms365 --schema 2>/dev/null; then
    echo ""
    echo "For detailed tool descriptions, see README.md"
else
    echo "Error: Could not list tools."
    echo ""
    echo "Possible solutions:"
    echo "  1. Ensure mcporter is installed and in your PATH"
    echo "  2. Check that ms-365-mcp-server is installed:"
    echo ""
    echo "       npm install -g @softeria/ms-365-mcp-server"
    echo ""
    echo "  3. Verify your mcporter.json configuration"
    exit 1
fi
