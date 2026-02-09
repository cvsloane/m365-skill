#!/usr/bin/env bash
# =============================================================================
# MS 365 Authentication Diagnostics
# =============================================================================
# This script checks the authentication status and environment configuration
# for the Microsoft 365 skill.
#
# Usage:
#   ./check-auth.sh
#
# It will report:
#   - Which environment variables are set
#   - Whether authentication is working
#   - Any obvious configuration issues
#
# Exit codes:
#   0 - Authentication successful
#   1 - Authentication failed or mcporter not available
# =============================================================================

set -euo pipefail

echo "MS 365 Authentication Diagnostics"
echo "================================="
echo ""

# Check environment variables
echo "Environment Variables:"
echo "----------------------"

if [[ -n "${MS365_MCP_CLIENT_ID:-}" ]]; then
    echo "  ✓ MS365_MCP_CLIENT_ID: Set (${MS365_MCP_CLIENT_ID:0:8}...)"
else
    echo "  ✗ MS365_MCP_CLIENT_ID: Not set"
fi

if [[ -n "${MS365_MCP_CLIENT_SECRET:-}" ]]; then
    echo "  ✓ MS365_MCP_CLIENT_SECRET: Set (***)"
else
    echo "  ✗ MS365_MCP_CLIENT_SECRET: Not set"
fi

if [[ -n "${MS365_MCP_TENANT_ID:-}" ]]; then
    echo "  ✓ MS365_MCP_TENANT_ID: ${MS365_MCP_TENANT_ID}"
else
    echo "  ✗ MS365_MCP_TENANT_ID: Not set (will use 'common')"
fi

echo ""

# Check optional flags
echo "Optional Configuration:"
echo "-----------------------"
echo "  MS365_MCP_ORG_MODE: ${MS365_MCP_ORG_MODE:-not set}"
echo "  MS365_MCP_READ_ONLY: ${MS365_MCP_READ_ONLY:-not set}"
echo "  MS365_MCP_OUTPUT_FORMAT: ${MS365_MCP_OUTPUT_FORMAT:-not set}"
echo ""

# Test authentication via mcporter
echo "Authentication Test:"
echo "--------------------"

if ! command -v mcporter &> /dev/null; then
    echo "  ✗ mcporter command not found"
    echo ""
    echo "  Make sure Clawdbot is properly installed and mcporter is in your PATH."
    exit 1
fi

# Try to list messages (limit 1 to be quick)
if mcporter call ms365.list_messages limit=1 &>/dev/null; then
    echo "  ✓ Authentication successful!"
    echo ""
    echo "  You can now use Microsoft 365 commands."
    exit 0
else
    echo "  ✗ Authentication failed"
    echo ""
    echo "  Possible solutions:"
    echo "    1. Run ./auth-device.sh for interactive authentication"
    echo "    2. Check your Azure AD credentials (if using app registration)"
    echo "    3. Verify the MS365_MCP_CLIENT_ID is correct"
    echo "    4. Ensure the client secret hasn't expired"
    echo ""
    exit 1
fi
