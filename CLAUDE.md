# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

## Repository Overview

This is a Clawdbot skill for Microsoft 365 integration via the Graph API. It uses the `@softeria/ms-365-mcp-server` MCP server through Clawdbot's `mcporter` bridge.

## Key Files

- **SKILL.md** - Main skill definition with Clawdbot frontmatter (required)
- **README.md** - User-facing setup guide with Azure AD instructions
- **scripts/** - Helper scripts for server management and auth

## Deployment Locations

- **Local dev**: `~/dev/m365-skill/`
- **apps-vps host**: `/data/clawdbot/workspace/skills/ms365/`
- **Container path**: `/root/clawd/skills/ms365/`

## Testing

The skill requires Azure AD credentials to function. Test with:

```bash
# Inside clawdbot container
mcporter list ms365
mcporter call ms365.list-mail-messages top=5
```

## Environment Variables

These must be set in Coolify for the clawdbot container:

| Variable | Description |
|----------|-------------|
| `MS365_MCP_CLIENT_ID` | Azure AD app client ID |
| `MS365_MCP_CLIENT_SECRET` | Azure AD app client secret |
| `MS365_MCP_TENANT_ID` | `consumers` for personal, or tenant ID |

## Related Documentation

- MCP Server: https://github.com/Softeria/ms-365-mcp-server
- Clawdbot Skills: See Clawdbot documentation
- Azure AD: https://portal.azure.com → App registrations

## Deployment Steps (for tmux session)

1. Clone to apps-vps:
   ```bash
   ssh root@apps-vps "cd /data/clawdbot/workspace/skills && git clone https://github.com/cvsloane/m365-skill ms365"
   ```

2. Add env vars in Coolify (clawdbot app → Environment Variables)

3. Update mcporter.json (see mcporter.example.json)

4. Restart clawdbot container in Coolify

5. Test:
   ```bash
   ssh root@apps-vps "docker exec clawdbot mcporter list ms365"
   ```
