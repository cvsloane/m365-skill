#!/usr/bin/env bash
# =============================================================================
# Start MS 365 MCP Server (HTTP Mode)
# =============================================================================
# This script starts the Microsoft 365 MCP server in HTTP mode, making it
# accessible to multiple clients or for debugging purposes.
#
# Usage:
#   ./start-server.sh
#
# Environment Variables:
#   MS365_MCP_PORT      - Port to listen on (default: 3365)
#   MS365_MCP_ORG_MODE  - Set to 'true' to enable Teams/SharePoint
#   MS365_MCP_READ_ONLY - Set to 'true' to disable write operations
#   MS365_MCP_OUTPUT_FORMAT - Set to 'toon' for compact output
#
# Example:
#   MS365_MCP_ORG_MODE=true ./start-server.sh
#
# Note: HTTP mode is useful for development but requires proper security
# considerations for production deployments.
# =============================================================================

set -euo pipefail

# Configuration
PORT="${MS365_MCP_PORT:-3365}"
ARGS=""

# Build argument list based on environment flags
if [[ "${MS365_MCP_ORG_MODE:-}" == "true" ]]; then
    ARGS="$ARGS --org-mode"
fi

if [[ "${MS365_MCP_READ_ONLY:-}" == "true" ]]; then
    ARGS="$ARGS --read-only"
fi

if [[ "${MS365_MCP_OUTPUT_FORMAT:-}" == "toon" ]]; then
    ARGS="$ARGS --toon"
fi

# Display startup information
echo "Starting MS 365 MCP Server"
echo "=========================="
echo "  Port:    $PORT"
echo "  Mode:    ${MS365_MCP_ORG_MODE:+org }${MS365_MCP_READ_ONLY:+read-only }standard"
echo "  Format:  ${MS365_MCP_OUTPUT_FORMAT:-default}"
echo "  Args:    $ARGS"
echo ""

# Start the server
# Using exec to replace shell process so signals are handled correctly
exec npx -y @softeria/ms-365-mcp-server --http "$PORT" $ARGS
