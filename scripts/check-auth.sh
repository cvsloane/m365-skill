#!/usr/bin/env bash
# Check MS 365 authentication status
set -euo pipefail

echo "Checking MS 365 authentication..."
echo ""

# Check environment variables
echo "Environment Variables:"
if [[ -n "${MS365_MCP_CLIENT_ID:-}" ]]; then
  echo "  MS365_MCP_CLIENT_ID: Set (${MS365_MCP_CLIENT_ID:0:8}...)"
else
  echo "  MS365_MCP_CLIENT_ID: Not set"
fi

if [[ -n "${MS365_MCP_CLIENT_SECRET:-}" ]]; then
  echo "  MS365_MCP_CLIENT_SECRET: Set (***)"
else
  echo "  MS365_MCP_CLIENT_SECRET: Not set"
fi

if [[ -n "${MS365_MCP_TENANT_ID:-}" ]]; then
  echo "  MS365_MCP_TENANT_ID: ${MS365_MCP_TENANT_ID}"
else
  echo "  MS365_MCP_TENANT_ID: Not set (will use 'common')"
fi

echo ""

# Try to list messages as auth test
echo "Testing authentication..."
if mcporter call ms365.list_messages top=1 2>/dev/null; then
  echo ""
  echo "Authentication successful!"
else
  echo ""
  echo "Authentication failed. Check credentials or run auth-device.sh"
fi
