# CLAUDE.md

> **Guidance for Claude Code when working with the m365-skill repository**

---

## Repository Overview

This repository contains a **Clawdbot skill** that enables Microsoft 365 integration through the Microsoft Graph API. It acts as a bridge between Clawdbot and the `@softeria/ms-365-mcp-server` MCP (Model Context Protocol) server.

### What This Skill Does

- Provides natural language access to Outlook email, Calendar, OneDrive, To Do, Contacts, and OneNote
- Offers both programmatic (CLI) and conversational interfaces
- Supports personal Microsoft accounts and organizational (work/school) accounts

### Architecture

```
┌─────────────┐     ┌──────────────┐     ┌──────────────────┐     ┌─────────────┐
│   Clawdbot  │────→│  mcporter    │────→│ @softeria/ms-365 │────→│ Microsoft   │
│   (User)    │←────│  (bridge)    │←────│ -mcp-server      │←────│ Graph API   │
└─────────────┘     └──────────────┘     └──────────────────┘     └─────────────┘
                              ↑
                              │
                    ┌─────────┴──────────┐
                    │  ms365_cli.py      │
                    │  (Python wrapper)  │
                    └────────────────────┘
```

---

## Key Files

| File | Purpose |
|------|---------|
| `SKILL.md` | **Primary skill definition** with Clawdbot frontmatter—defines activation triggers, commands, and behavior |
| `README.md` | User-facing documentation with installation guide, troubleshooting, and feature reference |
| `ms365_cli.py` | Python CLI wrapper that communicates with the MCP server via stdio |
| `mcporter.example.json` | Example configuration for the mcporter bridge |
| `scripts/` | Helper scripts for server management and authentication |

---

## Development Environment

### Local Development

```bash
# Repository location
~/dev/m365-skill/
```

### Deployment Locations

| Environment | Path |
|-------------|------|
| **apps-vps host** | `/data/clawdbot/workspace/skills/ms365/` |
| **Docker container** | `/root/clawd/skills/ms365/` |

### Prerequisites

- Python 3.6+ (standard library only—no pip dependencies)
- Node.js with npm
- Azure AD app registration (for testing)

---

## Testing & Verification

### Quick Verification

Test the MCP server connection:

```bash
# List available tools
mcporter list ms365

# Test email access
mcporter call ms365.list_messages limit=5

# Test calendar access
mcporter call ms365.list_events top=5
```

### Testing Inside the Clawdbot Container

```bash
# SSH to host, then execute in container
ssh root@apps-vps "docker exec clawdbot mcporter list ms365"
```

### Testing Authentication

```bash
# Check auth status
python3 ms365_cli.py status

# If not authenticated, use device flow
python3 ms365_cli.py login
```

---

## Environment Variables

These variables must be configured in Coolify for the clawdbot container:

| Variable | Required | Description |
|----------|----------|-------------|
| `MS365_MCP_CLIENT_ID` | **Yes** | Azure AD app client ID |
| `MS365_MCP_CLIENT_SECRET` | **Yes** | Azure AD app secret |
| `MS365_MCP_TENANT_ID` | **Yes** | Tenant ID (`consumers` for personal accounts) |
| `MS365_MCP_ORG_MODE` | No | Set to `true` to enable Teams/SharePoint |

### Where to Set These

1. Open Coolify dashboard
2. Navigate to the **clawdbot** application
3. Go to **Environment Variables**
4. Add the three required variables above
5. Restart the container

---

## Deployment Workflow

### Step 1: Clone to Production Host

```bash
ssh root@apps-vps "cd /data/clawdbot/workspace/skills && git clone https://github.com/cvsloane/m365-skill ms365"
```

### Step 2: Configure Environment Variables

Add the three required Azure credentials in Coolify (see above).

### Step 3: Configure mcporter

Copy the example configuration and customize:

```bash
# On the host
cp /data/clawdbot/workspace/skills/ms365/mcporter.example.json /data/clawdbot/mcporter.json

# Edit to add your actual credentials
nano /data/clawdbot/mcporter.json
```

### Step 4: Restart Clawdbot

Restart the container in Coolify to pick up new environment variables.

### Step 5: Verify Deployment

```bash
# Test from your local machine
ssh root@apps-vps "docker exec clawdbot mcporter list ms365"
```

Expected output: A list of available MS365 tools.

---

## Common Development Tasks

### Adding a New CLI Command

1. Add function in `ms365_cli.py` following existing pattern
2. Update `SKILL.md` with command documentation
3. Update `README.md` tool reference table if applicable
4. Test locally: `python3 ms365_cli.py <new_command>`

### Updating Documentation

- **User-facing changes** → `README.md`
- **Clawdbot behavior changes** → `SKILL.md`
- **Developer setup changes** → `CLAUDE.md`

### Troubleshooting Authentication Issues

```bash
# Clear cached tokens
rm -rf ~/.clawdbot/ms365-cache

# Re-authenticate
python3 ms365_cli.py login

# For headless environments, verify Azure app credentials
# are correctly set as environment variables
echo $MS365_MCP_CLIENT_ID
echo $MS365_MCP_TENANT_ID
```

---

## Related Resources

| Resource | URL |
|----------|-----|
| MCP Server Repository | https://github.com/Softeria/ms-365-mcp-server |
| Clawdbot Documentation | https://github.com/clawdbot/clawdbot |
| Azure Portal (App Registration) | https://portal.azure.com → Microsoft Entra ID → App registrations |
| Microsoft Graph API Docs | https://docs.microsoft.com/en-us/graph/api/overview |

---

## Notes

- The Python CLI wrapper uses **only standard library** modules—no virtualenv or pip required
- Authentication tokens are cached in `~/.clawdbot/ms365-cache/`
- The MCP server runs as a subprocess via stdio (not HTTP by default)
- For production deployments, always use Azure AD app registration—not device code flow
