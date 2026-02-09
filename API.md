# MS 365 Skill API Documentation

## Overview

The MS 365 Skill provides programmatic access to Microsoft 365 services through the Graph API via the `@softeria/ms-365-mcp-server` MCP server. This document describes the available tools and their usage patterns.

## Authentication

### Device Code Flow (Interactive)
```bash
python3 ms365_cli.py login
```

### Azure AD App (Headless)
Set environment variables:
- `MS365_MCP_CLIENT_ID` - Azure AD app client ID
- `MS365_MCP_CLIENT_SECRET` - Azure AD app client secret  
- `MS365_MCP_TENANT_ID` - Tenant ID or "consumers"

## Available Tools

### Email (Outlook) Tools

#### `list-mail-messages`
List emails in a folder.

**Parameters:**
- `top` (number, optional): Maximum number of messages to return
- `folderId` (string, optional): Folder ID (default: "Inbox")

**Example:**
```bash
mcporter call ms365.list-mail-messages top=10
```

#### `get-mail-message`
Get full email content.

**Parameters:**
- `messageId` (string, required): Message ID

**Example:**
```bash
mcporter call ms365.get-mail-message messageId="AAMkAGI...="
```

#### `send-mail`
Send a new email.

**Parameters:**
- `body` (object): Email message structure
  - `message`: Message object
    - `subject` (string): Email subject
    - `body` (object): Message body
      - `contentType` (string): "Text" or "HTML"
      - `content` (string): Message content
    - `toRecipients` (array): Recipients
      - `emailAddress` (object): Email address
        - `address` (string): Email address

**Example:**
```bash
mcporter call ms365.send-mail '{
  "message": {
    "subject": "Test Email",
    "body": {
      "contentType": "Text",
      "content": "Hello World"
    },
    "toRecipients": [
      {"emailAddress": {"address": "recipient@example.com"}}
    ]
  }
}'
```

### Calendar Tools

#### `list-calendar-events`
List calendar events.

**Parameters:**
- `top` (number, optional): Maximum number of events to return

**Example:**
```bash
mcporter call ms365.list-calendar-events top=5
```

#### `create-calendar-event`
Create a new calendar event.

**Parameters:**
- `body` (object): Event structure
  - `subject` (string): Event title
  - `start` (object): Start time
    - `dateTime` (string): ISO 8601 datetime
    - `timeZone` (string): Time zone (default: "America/Chicago")
  - `end` (object): End time
    - `dateTime` (string): ISO 8601 datetime
    - `timeZone` (string): Time zone (default: "America/Chicago")
  - `body` (object, optional): Event description
    - `contentType` (string): "Text" or "HTML"
    - `content` (string): Description content

**Example:**
```bash
mcporter call ms365.create-calendar-event '{
  "subject": "Team Meeting",
  "start": {
    "dateTime": "2026-01-15T10:00:00",
    "timeZone": "America/Chicago"
  },
  "end": {
    "dateTime": "2026-01-15T11:00:00",
    "timeZone": "America/Chicago"
  },
  "body": {
    "contentType": "Text",
    "content": "Weekly team sync"
  }
}'
```

### OneDrive Tools

#### `list-folder-files`
List files in a folder.

**Parameters:**
- `driveId` (string): Drive ID ("me" for personal OneDrive)
- `driveItemId` (string): Folder ID ("root" for root folder)

**Example:**
```bash
mcporter call ms365.list-folder-files '{
  "driveId": "me",
  "driveItemId": "root"
}'
```

### To Do Tools

#### `list-todo-task-lists`
List all task lists.

**Parameters:**
- None

**Example:**
```bash
mcporter call ms365.list-todo-task-lists
```

#### `list-todo-tasks`
List tasks from a specific list.

**Parameters:**
- `todoTaskListId` (string, required): Task list ID

**Example:**
```bash
mcporter call ms365.list-todo-tasks '{
  "todoTaskListId": "AAMkAGI...="
}'
```

#### `create-todo-task`
Create a new task.

**Parameters:**
- `todoTaskListId` (string, required): Task list ID
- `body` (object): Task structure
  - `title` (string): Task title
  - `dueDateTime` (object, optional): Due date
    - `dateTime` (string): ISO 8601 datetime
    - `timeZone` (string): Time zone

**Example:**
```bash
mcporter call ms365.create-todo-task '{
  "todoTaskListId": "AAMkAGI...=",
  "body": {
    "title": "Review budget",
    "dueDateTime": {
      "dateTime": "2026-01-20T17:00:00",
      "timeZone": "America/Chicago"
    }
  }
}'
```

### Contact Tools

#### `list-outlook-contacts`
List contacts.

**Parameters:**
- `top` (number, optional): Maximum number of contacts to return

**Example:**
```bash
mcporter call ms365.list-outlook-contacts top=20
```

#### `search-people`
Search contacts.

**Parameters:**
- `search` (string, required): Search query

**Example:**
```bash
mcporter call ms365.search-people '{
  "search": "John"
}'
```

### Organization Tools (--org-mode)

These tools are only available when `MS365_MCP_ORG_MODE=true` is set.

#### `list-teams`
List Microsoft Teams.

**Parameters:**
- None

**Example:**
```bash
mcporter call ms365.list-teams
```

#### `send-channel-message`
Post a message to a Teams channel.

**Parameters:**
- `teamId` (string, required): Team ID
- `channelId` (string, required): Channel ID
- `message` (string, required): Message content

**Example:**
```bash
mcporter call ms365.send-channel-message '{
  "teamId": "19:...@thread.skype",
  "channelId": "19:...@thread.skype",
  "message": "Hello team!"
}'
```

## Error Handling

Common error responses:

```json
{
  "error": {
    "code": "AuthenticationError",
    "message": "Authentication failed"
  }
}
```

```json
{
  "error": {
    "code": "PermissionDenied",
    "message": "Insufficient privileges"
  }
}
```

## Configuration Options

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `MS365_MCP_CLIENT_ID` | Azure AD app client ID | Required for headless |
| `MS365_MCP_CLIENT_SECRET` | Azure AD app client secret | Required for headless |
| `MS365_MCP_TENANT_ID` | Tenant ID | "common" |
| `MS365_MCP_ORG_MODE` | Enable organization features | false |
| `MS365_MCP_READ_ONLY` | Disable write operations | false |
| `MS365_MCP_OUTPUT_FORMAT` | Output format ("toon") | standard |
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

## Rate Limits

The Microsoft Graph API has rate limits. Typical limits:
- 60 requests per minute per application
- 10 requests per second per user

## Troubleshooting

### Common Issues

1. **Authentication errors**: Check environment variables and Azure AD permissions
2. **Permission denied**: Verify required API permissions are granted
3. **Rate limiting**: Implement retry logic with exponential backoff
4. **Timeout**: Increase timeout for large operations

### Debug Commands

```bash
# Check authentication status
python3 ms365_cli.py status

# List available tools
mcporter list ms365

# Test with minimal data
mcporter call ms365.list-mail-messages top=1
```