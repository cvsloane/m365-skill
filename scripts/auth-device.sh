#!/usr/bin/env bash
# Interactive device code authentication for MS 365
set -euo pipefail

echo "Starting MS 365 device code authentication..."
echo ""
echo "You will be prompted to:"
echo "  1. Open https://microsoft.com/devicelogin in a browser"
echo "  2. Enter the code displayed"
echo "  3. Sign in with your Microsoft account"
echo ""

exec npx -y @softeria/ms-365-mcp-server --auth-only
