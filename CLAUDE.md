# CLAUDE.md — Developer Guide for MS 365 Skill

> Quick reference for Claude Code and other AI assistants working with this repository.

---

## Repository Overview

This is a **Clawdbot skill** that integrates Microsoft 365 services (Outlook, Calendar, OneDrive, Teams, etc.) via the Microsoft Graph API. It wraps the `@softeria/ms-365-mcp-server` package and exposes functionality through Clawdbot's `mcporter` bridge.

### Architecture

```
┌─────────────┐     ┌────────────────┐     ┌─────────────────────┐
│  Clawdbot   │────▶│   mcporter     │────▶│ ms-365-mcp-server   │
│  (Natural   │     │   (bridge)     │     │ (Microsoft Graph    │
│   language) │◀────│                │◀────│  API wrapper)       │
└─────────────┘     └────────────────┘     └─────────────────────┘
                            │
                            ▼
                   ┌────────────────┐
                   │  ms365_cli.py  │
                   │  (CLI wrapper) │
                   └────────────────┘
```

### Key Files

| File | Purpose |
|------|---------|
| `SKILL.md` | Skill definition with Clawdbot frontmatter (required) |
| `SKILL.clawdbot.md` | Identical copy for legacy compatibility |
| `README.md` | User-facing installation and usage guide |
| `ms365_cli.py` | Python CLI wrapper around the MCP server |
| `scripts/*.sh` | Helper scripts for auth, diagnostics, and server management |
| `mcporter.example.json` | Example mcporter configuration |

---

## Deployment Locations

Know where the skill lives in different environments:

| Environment | Path |
|-------------|------|
| Local development | `~/dev/m365-skill/` |
| apps-vps host | `/data/clawdbot/workspace/skills/ms365/` |
| Clawdbot container | `/root/clawd/skills/ms365/` |

---

## Environment Variables

These must be configured in your deployment environment (Coolify, Docker, or `.env`):

| Variable | Required For | Description |
|----------|--------------|-------------|
| `MS365_MCP_CLIENT_ID` | Azure AD auth | Azure AD application (client) ID |
| `MS365_MCP_CLIENT_SECRET` | Azure AD auth | Client secret from Azure portal |
| `MS365_MCP_TENANT_ID` | Azure AD auth | `consumers` for personal accounts, or tenant GUID |
| `MS365_MCP_ORG_MODE` | Teams/SharePoint | Set to `true` to enable organization features |
| `MS365_MCP_READ_ONLY` | Optional | Set to `true` to disable write operations |
| `MS365_MCP_OUTPUT_FORMAT` | Optional | Set to `toon` for token-efficient output |

### Coolify Configuration

For the Clawdbot deployment on apps-vps, add these in Coolify → Applications → clawdbot → Environment Variables.

---

## Testing & Debugging

### Basic Connectivity

```bash
# List available tools
mcporter list ms365

# Test email access
mcporter call ms365.list_messages limit=5

# Test calendar
mcporter call ms365.list_events top=5
```

### Authentication Diagnostics

```bash
# Check auth status via CLI wrapper
python3 ms365_cli.py status

# List cached accounts
python3 ms365_cli.py accounts

# Interactive re-auth
scripts/auth-device.sh
```

### Direct Server Testing

```bash
# Start server in HTTP mode for debugging
scripts/start-server.sh

# Test via HTTP
curl http://localhost:3365/health  # (if available)
```

### Common Debug Commands

```bash
# Verify npm package installation
npm list -g @softeria/ms-365-mcp-server

# Run server directly with verbose output
DEBUG=* npx -y @softeria/ms-365-mcp-server

# Check environment variables
scripts/check-auth.sh
```

---

## Deployment Workflow

### Initial Deployment (apps-vps)

```bash
# 1. Clone to server
ssh root@apps-vps "cd /data/clawdbot/workspace/skills && \
  git clone https://github.com/cvsloane/m365-skill ms365"

# 2. Configure environment variables in Coolify
#    - MS365_MCP_CLIENT_ID
#    - MS365_MCP_CLIENT_SECRET
#    - MS365_MCP_TENANT_ID

# 3. Update mcporter.json (see mcporter.example.json)

# 4. Restart container in Coolify UI

# 5. Verify deployment
ssh root@apps-vps "docker exec clawdbot mcporter list ms365"
```

### Updates

```bash
# Pull latest changes
ssh root@apps-vps "cd /data/clawdbot/workspace/skills/ms365 && git pull"

# Restart to pick up changes
ssh root@apps-vps "docker restart clawdbot"
```

---

## Modifying the Code

### ms365_cli.py Structure

The CLI is organized by service with command handlers:

```python
# Authentication
cmd_login(), cmd_status(), cmd_accounts(), cmd_user()

# Email
cmd_mail_list(), cmd_mail_read(), cmd_mail_send()

# Calendar
cmd_calendar_list(), cmd_calendar_create()

# OneDrive
cmd_files_list()

# Tasks
cmd_tasks_list(), cmd_tasks_get(), cmd_tasks_create()

# Contacts
cmd_contacts_list(), cmd_contacts_search()
```

### Adding a New Command

1. **Define the command handler:**
   ```python
   def cmd_new_feature(args):
       result = call_mcp("new-feature-tool", {"param": args.value})
       format_output(result)
   ```

2. **Add argument parser:**
   ```python
   new_p = subparsers.add_parser('newcmd', help='Description')
   new_p.add_argument('--value', required=True)
   new_p.set_defaults(func=cmd_new_feature)
   ```

3. **Update SKILL.md** with new command documentation

### call_mcp() Function

This is the core wrapper that communicates with the MCP server via stdio:

```python
def call_mcp(method: str, params: dict = None) -> dict:
    """
    Call an MCP tool method.
    
    Args:
        method: The MCP tool name (e.g., "list-mail-messages")
        params: Dictionary of tool arguments
    
    Returns:
        Parsed JSON response or error dict
    """
```

---

## Troubleshooting Common Issues

### "Application not found" (AADSTS700016)

- Verify `MS365_MCP_CLIENT_ID` matches Azure portal
- Check tenant ID is correct
- Ensure app registration hasn't been deleted

### "Invalid client secret" (AADSTS7000218)

- Secret may have expired—check expiration date in Azure
- Recreate secret if needed
- Ensure no trailing whitespace in env var

### "Insufficient privileges" (403)

- Missing API permissions in Azure AD
- Admin consent not granted for organization features
- Check `MS365_MCP_ORG_MODE` if using Teams/SharePoint

### "Token expired" (401)

- Re-authenticate: `scripts/auth-device.sh`
- Clear token cache and retry
- Check system clock accuracy

### mcporter connection failures

- Verify mcporter.json syntax: `python3 -m json.tool ~/.clawdbot/mcporter.json`
- Check npm package: `npm list -g @softeria/ms-365-mcp-server`
- Test server directly: `npx -y @softeria/ms-365-mcp-server --help`

---

## Related Documentation

| Resource | URL |
|----------|-----|
| MCP Server (Softeria) | https://github.com/Softeria/ms-365-mcp-server |
| Microsoft Graph API | https://docs.microsoft.com/en-us/graph/ |
| Azure Portal | https://portal.azure.com |
| Clawdbot Docs | (internal) |

---

## Quick Command Reference

| Task | Command |
|------|---------|
| List tools | `mcporter list ms365` |
| Get email | `mcporter call ms365.list_messages limit=5` |
| Check auth | `python3 ms365_cli.py status` |
| Re-authenticate | `scripts/auth-device.sh` |
| Start HTTP server | `scripts/start-server.sh` |
| Diagnose issues | `scripts/check-auth.sh` |

---

<p align="center"><em>Last updated for skill version compatible with @softeria/ms-365-mcp-server latest</em></p>
