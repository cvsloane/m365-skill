# MS 365 Skill API Documentation

## Overview

This skill provides a comprehensive interface to Microsoft 365 services through the Microsoft Graph API via the `@softeria/ms-365-mcp-server` MCP server.

## Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   User/Agent    │    │   Clawbot       │    │   MS 365        │
│                 │    │                 │    │   Services      │
│ - Natural Lang  │◄──►│ - mcporter      │◄──►│ - Outlook       │
│ - CLI Interface │    │ - Skill System  │    │ - OneDrive      │
│                 │    │                 │    │ - Calendar      │
└─────────────────┘    └─────────────────┘    │ - To Do         │
                                              │ - Teams         │
                                              │ - SharePoint    │
                                              └─────────────────┘
                                                    ↑
┌─────────────────────────────────────────────────────────────┐
│              @softeria/ms-365-mcp-server                   │
│              (Microsoft Graph API Client)                   │
└─────────────────────────────────────────────────────────────┘
```

## Authentication Methods

### 1. Azure AD App Registration (Recommended)
- **Purpose**: Headless operation, Docker deployments
- **Environment Variables**:
  - `MS365_MCP_CLIENT_ID`: Azure AD app client ID
  - `MS365_MCP_CLIENT_SECRET`: Azure AD app client secret  
  - `MS365_MCP_TENANT_ID`: Tenant ID or "consumers"

### 2. Device Code Flow
- **Purpose**: Interactive testing, development
- **Usage**: Run `python3 ms365_cli.py login` or `scripts/auth-device.sh`

## Core Components

### 1. MCP Server Interface
The skill communicates with Microsoft 365 through the MCP server protocol:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "tool_name",
    "arguments": {...}
  }
}
```

### 2. Python CLI Wrapper (`ms365_cli.py`)
- **Purpose**: Direct command-line access to MS 365 services
- **Location**: `/ms365_cli.py`
- **Usage**: `python3 ms365_cli.py <command> [options]`

### 3. mcporter Integration
- **Purpose**: Bridge between Clawbot and MCP server
- **Configuration**: `~/.clawdbot/mcporter.json`
- **Usage**: `mcporter call ms365.<tool_name> [parameters]`

## Available Tools by Category

### Email Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-mail-messages` | `top`, `folderId` | List emails in folder |
| `get-mail-message` | `messageId` | Get specific email |
| `send-mail` | `body` (message object) | Send new email |
| `reply-mail-message` | `messageId`, `body` | Reply to email |
| `forward-mail-message` | `messageId`, `body` | Forward email |
| `delete-mail-message` | `messageId` | Delete email |
| `move-mail-message` | `messageId`, `destinationId` | Move to folder |
| `search-mail-messages` | `query` | Search emails |

### Calendar Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-calendar-events` | `top`, `startDateTime`, `endDateTime` | List events |
| `get-calendar-event` | `eventId` | Get specific event |
| `create-calendar-event` | `body` (event object) | Create event |
| `update-calendar-event` | `eventId`, `body` | Update event |
| `delete-calendar-event` | `eventId` | Delete event |
| `list-calendars` | | List all calendars |

### OneDrive Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-folder-files` | `driveId`, `driveItemId`, `top` | List files/folders |
| `get-file` | `driveId`, `fileId` | Get file metadata |
| `download-file` | `driveId`, `fileId` | Download file content |
| `upload-file` | `driveId`, `parentFolderId`, `filename`, `content` | Upload file |
| `create-folder` | `driveId`, `parentFolderId`, `name` | Create folder |
| `delete-file` | `driveId`, `fileId` | Delete file/folder |
| `search-files` | `query` | Search files |

### To Do Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-todo-task-lists` | | List task lists |
| `list-todo-tasks` | `todoTaskListId`, `top` | List tasks in list |
| `get-todo-task` | `todoTaskListId`, `taskId` | Get specific task |
| `create-todo-task` | `todoTaskListId`, `body` | Create task |
| `update-todo-task` | `todoTaskListId`, `taskId`, `body` | Update task |
| `complete-todo-task` | `todoTaskListId`, `taskId` | Mark complete |
| `delete-todo-task` | `todoTaskListId`, `taskId` | Delete task |

### Contact Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-outlook-contacts` | `top` | List contacts |
| `get-outlook-contact` | `contactId` | Get specific contact |
| `search-people` | `search` | Search contacts |
| `create-outlook-contact` | `body` | Create contact |

### OneNote Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-notebooks` | | List notebooks |
| `list-notebook-sections` | `notebookId` | List sections |
| `list-section-pages` | `sectionId` | List pages |
| `get-page-content` | `pageId` | Get page content |

### Organization Tools (Org Mode)
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-teams` | | List Teams |
| `list-team-channels` | `teamId` | List channels |
| `send-channel-message` | `teamId`, `channelId`, `message` | Post to channel |
| `list-chats` | | List chats |
| `send-chat-message` | `chatId`, `message` | Send chat message |
| `list-sites` | | List SharePoint sites |
| `list-site-files` | `siteId` | List site documents |

## Configuration Options

### Environment Variables
| Variable | Description | Default |
|----------|-------------|---------|
| `MS365_MCP_CLIENT_ID` | Azure AD app client ID | Required for headless |
| `MS365_MCP_CLIENT_SECRET` | Azure AD app client secret | Required for headless |
| `MS365_MCP_TENANT_ID` | Tenant ID | "common" |
| `MS365_MCP_ORG_MODE` | Enable org features | false |
| `MS365_MCP_READ_ONLY` | Read-only mode | false |
| `MS365_MCP_OUTPUT_FORMAT` | Output format (toon) | standard |
| `MS365_MCP_PORT` | HTTP server port | 3365 |

### mcporter Configuration
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

## Error Handling

### Common Error Codes
- `AADSTS700016`: Application not found
- `AADSTS7000218`: Invalid client secret
- `AADSTS7000218`: Invalid client secret
- `401 Unauthorized`: Token expired or invalid
- `403 Forbidden`: Insufficient privileges

### Troubleshooting Steps
1. Check environment variables
2. Verify Azure AD app permissions
3. Test authentication with `mcporter call ms365.verify-login`
4. Check token expiration and re-authenticate if needed

## Development

### Adding New Tools
1. Check MCP server capabilities: `mcporter list ms365 --schema`
2. Add corresponding CLI command in `ms365_cli.py`
3. Update documentation in README.md and SKILL.md
4. Test with both CLI and mcporter interfaces

### Testing
```bash
# Test authentication
python3 ms365_cli.py status

# Test email access
mcporter call ms365.list_messages limit=5

# Test calendar access
mcporter call ms365.list_events top=5
```

## Security Considerations

- Store client secrets securely (environment variables, not in code)
- Use appropriate OAuth scopes for your use case
- Regularly rotate client secrets
- Enable conditional access policies for organizational accounts
- Consider read-only mode for sensitive operations

## Performance Optimization

- Use `--toon` format to reduce token usage by 30-60%
- Implement pagination with `top` parameter for large datasets
- Cache frequently accessed data where appropriate
- Use discovery mode to load only needed tools

## Integration Examples

### Clawbot Natural Language
```
User: "Check my email"
Agent: Runs `mcporter call ms365.list_messages limit=10`

User: "Schedule meeting tomorrow at 2pm"
Agent: Runs `mcporter call ms365.create-calendar-event` with appropriate parameters
```

### Direct CLI Usage
```bash
# List recent emails
python3 ms365_cli.py mail list --top 5

# Create calendar event
python3 ms365_cli.py calendar create \
  --subject "Team Meeting" \
  --start "2026-02-20T14:00:00" \
  --end "2026-02-20T15:00:00"
```

### Programmatic Access
```python
import subprocess
import json

def call_ms365_tool(tool_name, params=None):
    """Call MS 365 tool programmatically"""
    # Implementation similar to ms365_cli.py call_mcp function
    pass
```