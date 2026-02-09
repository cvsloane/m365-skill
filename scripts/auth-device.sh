#!/usr/bin/env bash
# =============================================================================
# MS 365 Device Code Authentication
# =============================================================================
# This script initiates the device code flow for Microsoft 365 authentication.
# Use this when you want to authenticate interactively without setting up
# an Azure AD application.
#
# Usage:
#   ./auth-device.sh
#
# The script will:
#   1. Generate a device code
#   2. Display instructions to visit https://microsoft.com/devicelogin
#   3. Wait for you to complete authentication in your browser
#   4. Cache credentials for future use
#
# Note: This method requires interactive access and won't work in headless
# Docker containers. For automated deployments, use Azure AD app credentials.
# =============================================================================

set -euo pipefail

echo "Starting MS 365 device code authentication..."
echo ""
echo "You will be prompted to:"
echo "  1. Open https://microsoft.com/devicelogin in a browser"
echo "  2. Enter the code displayed below"
echo "  3. Sign in with your Microsoft account"
echo ""

# Execute the MCP server's built-in device code authentication
exec npx -y @softeria/ms-365-mcp-server --auth-only
