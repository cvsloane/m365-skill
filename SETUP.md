# MS 365 Skill Environment Setup Guide

This guide provides comprehensive setup instructions for the MS 365 Skill across different environments.

## Prerequisites

### System Requirements
- **Python**: 3.6 or higher (for CLI tool)
- **Node.js**: 14 or higher (for MCP server)
- **npm**: 6 or higher (for package management)
- **Git**: For repository cloning

### Network Requirements
- Internet connection to Microsoft Graph API
- Access to Azure Portal (for Azure AD setup)
- Firewall access to `graph.microsoft.com` endpoints

## Installation Steps

### 1. Clone the Repository

```bash
# Clone to your preferred location
git clone https://github.com/cvsloane/m365-skill.git
cd m365-skill

# Or for Clawbot skills directory
cd ~/.clawdbot/skills
git clone https://github.com/cvsloane/m365-skill ms365
```

### 2. Install Dependencies

#### MCP Server Installation
```bash
# Install globally (recommended)
npm install -g @softeria/ms-365-mcp-server

# Or install locally
npm install @softeria/ms-365-mcp-server
```

#### Python Dependencies
```bash
# The CLI tool uses only standard library - no additional packages needed
# Verify Python 3.6+ is available
python3 --version
```

### 3. Authentication Setup

Choose one of the following authentication methods:

#### Option A: Azure AD App Registration (Recommended for Production)

**Step 1: Create Azure AD App**
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: `Clawbot MS365` (or your choice)
   - **Supported account types**: Choose appropriate option
   - **Redirect URI**: Select "Web" and enter `http://localhost:3365/callback`

**Step 2: Configure API Permissions**
1. Go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add required permissions:
   - `Mail.ReadWrite` - Email access
   - `Calendars.ReadWrite` - Calendar access
   - `Files.ReadWrite` - OneDrive access
   - `Tasks.ReadWrite` - To Do access
   - `Contacts.Read` - Contacts access
   - `User.Read` - Basic profile
   - `Notes.Read` - OneNote access

**Step 4: Create Client Secret**
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description and expiration
4. **Copy the secret value immediately** (shown only once!)

**Step 5: Note Credentials**
- **Application (client) ID**: `MS365_MCP_CLIENT_ID`
- **Directory (tenant) ID**: `MS365_MCP_TENANT_ID`
- **Client secret**: `MS365_MCP_CLIENT_SECRET`

#### Option B: Device Code Flow (For Development/Testing)

No Azure setup required. Use device code authentication when needed.

### 4. Environment Configuration

#### Set Environment Variables
```bash
# For bash/zsh
export MS365_MCP_CLIENT_ID="your-client-id"
export MS365_MCP_CLIENT_SECRET="your-client-secret"
export MS365_MCP_TENANT_ID="your-tenant-id"

# For fish
set -x MS365_MCP_CLIENT_ID "your-client-id"
set -x MS365_MCP_CLIENT_SECRET "your-client-secret"
set -x MS365_MCP_TENANT_ID "your-tenant-id"

# Add to ~/.bashrc or ~/.zshrc for persistence
echo 'export MS365_MCP_CLIENT_ID="your-client-id"' >> ~/.bashrc
echo 'export MS365_MCP_CLIENT_SECRET="your-client-secret"' >> ~/.bashrc
echo 'export MS365_MCP_TENANT_ID="your-tenant-id"' >> ~/.bashrc
```

#### Create mcporter Configuration
```bash
# Copy example configuration
cp mcporter.example.json ~/.clawdbot/mcporter.json

# Edit the configuration
nano ~/.clawdbot/mcporter.json
```

Example configuration:
```json
{
  "servers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server"],
      "env": {
        "MS365_MCP_CLIENT_ID": "${MS365_MCP_CLIENT_ID}",
        "MS365_MCP_CLIENT_SECRET": "${MS365_MCP_CLIENT_SECRET}",
        "MS365_MCP_TENANT_ID": "${MS365_MCP_TENANT_ID}"
      }
    }
  }
}
```

### 5. Verify Installation

#### Test Authentication
```bash
# Check if server is accessible
mcporter list ms365

# Test basic functionality
mcporter call ms365.verify-login

# Test email access
mcporter call ms365.list-mail-messages limit=5

# Test calendar access
mcporter call ms365.list-calendar-events top=5
```

#### Test CLI Tool
```bash
# Check authentication status
python3 ms365_cli.py status

# List recent emails
python3 ms365_cli.py mail list --top 5

# List calendar events
python3 ms365_cli.py calendar list --top 5
```

## Environment-Specific Setup

### Development Environment
```bash
# Use device code for development
export MS365_MCP_USE_DEVICE_CODE=true

# Enable debug logging
export MS365_MCP_DEBUG=true

# Test with CLI tool
python3 ms365_cli.py login
python3 ms365_cli.py mail list --top 10
```

### Production Environment (Docker)
```bash
# Docker environment variables
docker run -e MS365_MCP_CLIENT_ID="$CLIENT_ID" \
           -e MS365_MCP_CLIENT_SECRET="$CLIENT_SECRET" \
           -e MS365_MCP_TENANT_ID="$TENANT_ID" \
           clawbot-image

# Or use docker-compose
version: '3.8'
services:
  clawbot:
    environment:
      - MS365_MCP_CLIENT_ID=${MS365_MCP_CLIENT_ID}
      - MS365_MCP_CLIENT_SECRET=${MS365_MCP_CLIENT_SECRET}
      - MS365_MCP_TENANT_ID=${MS365_MCP_TENANT_ID}
```

### Cloud Deployment (Coolify/Heroku)
```bash
# Set environment variables in your deployment platform
MS365_MCP_CLIENT_ID=your-client-id
MS365_MCP_CLIENT_SECRET=your-client-secret
MS365_MCP_TENANT_ID=your-tenant-id

# Optional: Enable organization features
MS365_MCP_ORG_MODE=true

# Optional: Read-only mode
MS365_MCP_READ_ONLY=true
```

## Advanced Configuration

### Server Mode Configuration
```json
{
  "servers": {
    "ms365": {
      "url": "http://localhost:3365",
      "env": {
        "MS365_MCP_CLIENT_ID": "${MS365_MCP_CLIENT_ID}",
        "MS365_MCP_CLIENT_SECRET": "${MS365_MCP_CLIENT_SECRET}",
        "MS365_MCP_TENANT_ID": "${MS365_MCP_TENANT_ID}"
      }
    }
  }
}
```

### Custom Port Configuration
```bash
export MS365_MCP_PORT=8080

# Start server on custom port
./scripts/start-server.sh
```

### Output Format Optimization
```bash
# Enable TOON format for reduced token usage
export MS365_MCP_OUTPUT_FORMAT=toon

# Configure mcporter for TOON format
{
  "servers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--toon"],
      "env": {
        "MS365_MCP_CLIENT_ID": "${MS365_MCP_CLIENT_ID}",
        "MS365_MCP_CLIENT_SECRET": "${MS365_MCP_CLIENT_SECRET}",
        "MS365_MCP_TENANT_ID": "${MS365_MCP_TENANT_ID}"
      }
    }
  }
}
```

## Troubleshooting Setup Issues

### Common Issues and Solutions

#### Authentication Errors
```bash
# Check environment variables
echo $MS365_MCP_CLIENT_ID
echo $MS365_MCP_CLIENT_SECRET
echo $MS365_MCP_TENANT_ID

# Test authentication
python3 ms365_cli.py status

# Re-authenticate if needed
python3 ms365_cli.py login
```

#### MCP Server Issues
```bash
# Check if MCP server is installed
npm list -g @softeria/ms-365-mcp-server

# Test MCP server directly
npx -y @softeria/ms-365-mcp-server --help

# Check server logs
./scripts/check-auth.sh
```

#### mcporter Configuration Issues
```bash
# Validate mcporter configuration
cat ~/.clawdbot/mcporter.json | python3 -m json.tool

# Test mcporter connection
mcporter list ms365

# Test specific tool
mcporter call ms365.list-mail-messages limit=1
```

### Debug Mode
```bash
# Enable debug logging
export MS365_MCP_DEBUG=true

# Run with debug output
python3 ms365_cli.py status 2>&1 | tee debug.log

# Check MCP server logs
npx -y @softeria/ms-365-mcp-server --debug
```

## Testing Your Setup

### Functional Tests
```bash
# Test all major features
python3 ms365_cli.py mail list --top 3
python3 ms365_cli.py calendar list --top 3
python3 ms365_cli.py files list
python3 ms365_cli.py tasks lists
python3 ms365_cli.py contacts list --top 3
```

### Integration Tests
```bash
# Test with mcporter
mcporter call ms365.list-mail-messages limit=5
mcporter call ms365.list-calendar-events top=5
mcporter call ms365.verify-login
```

### Performance Tests
```bash
# Test with larger datasets
python3 ms365_cli.py mail list --top 50
python3 ms365_cli.py calendar list --top 20
python3 ms365_cli.py contacts list --top 50
```

## Next Steps

1. **Review Documentation**: Read [README.md](./README.md) for usage instructions
2. **Explore API**: Check [API.md](./API.md) for detailed API documentation
3. **Understand Architecture**: Review [ARCHITECTURE.md](./ARCHITECTURE.md)
4. **Start Using**: Begin with basic email and calendar operations
5. **Customize**: Modify configuration for your specific needs

## Support

If you encounter issues during setup:
1. Check the troubleshooting section above
2. Review the [README.md](./README.md) troubleshooting guide
3. Check the MCP server documentation: https://github.com/Softeria/ms-365-mcp-server
4. Open an issue on GitHub if problems persist