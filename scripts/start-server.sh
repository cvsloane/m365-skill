#!/usr/bin/env bash
# Start MS 365 MCP server in HTTP mode
set -euo pipefail

PORT="${MS365_MCP_PORT:-3365}"
ARGS=""

# Add org mode if enabled
if [[ "${MS365_MCP_ORG_MODE:-}" == "true" ]]; then
  ARGS="$ARGS --org-mode"
fi

# Add read-only if enabled
if [[ "${MS365_MCP_READ_ONLY:-}" == "true" ]]; then
  ARGS="$ARGS --read-only"
fi

# Add TOON format if enabled
if [[ "${MS365_MCP_OUTPUT_FORMAT:-}" == "toon" ]]; then
  ARGS="$ARGS --toon"
fi

echo "Starting MS 365 MCP server..."
echo "  Port: $PORT"
echo "  Args: $ARGS"
echo ""

exec npx -y @softeria/ms-365-mcp-server --http "$PORT" $ARGS
