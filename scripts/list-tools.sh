#!/usr/bin/env bash
# List all available MS 365 tools
set -euo pipefail

echo "Available MS 365 Tools"
echo "======================"
echo ""

mcporter list ms365 --schema 2>/dev/null || {
  echo "Error: Could not list tools."
  echo "Make sure ms-365-mcp-server is installed:"
  echo "  npm install -g @softeria/ms-365-mcp-server"
  exit 1
}
