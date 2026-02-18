# MS 365 Skill API Documentation

## Overview

This skill provides access to Microsoft 365 services through the Microsoft Graph API via the `@softeria/ms-365-mcp-server` MCP server. All interactions are mediated through Clawbot's `mcporter` bridge.

## Authentication

### Device Code Flow (Interactive)
```bash
python3 ms365_cli.py login
```

### Azure AD App (Headless)
Set environment variables:
- `MS365_MCP_CLIENT_ID`
- `MS365_MCP_CLIENT_SECRET` 
- `MS365_MCP_TENANT_ID`

## Available Tools

### Email (Outlook)
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-mail-messages` | `top`, `folderId` | List emails in folder |
| `get-mail-message` | `messageId` | Get email details |
| `send-mail` | `body` (message object) | Send new email |
| `reply-mail-message` | `messageId`, `body` | Reply to email |
| `forward-mail-message` | `messageId`, `body` | Forward email |
| `delete-mail-message` | `messageId` | Delete email |
| `move-mail-message` | `messageId`, `destinationId` | Move to folder |
| `search-mail-messages` | `query` | Search emails |

### Calendar
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-calendar-events` | `top` | List calendar events |
| `get-calendar-event` | `eventId` | Get event details |
| `create-calendar-event` | `body` (event object) | Create event |
| `update-calendar-event` | `eventId`, `body` | Update event |
| `delete-calendar-event` | `eventId` | Delete event |
| `list-calendars` | | List all calendars |

### OneDrive
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-folder-files` | `driveId`, `driveItemId` | List files/folders |
| `get-file-metadata` | `driveId`, `itemId` | Get file details |
| `download-file-content` | `driveId`, `itemId` | Download file |
| `upload-file` | `driveId`, `parentItemId`, `file` | Upload file |
| `create-folder` | `driveId`, `name`, `parentItemId` | Create folder |
| `delete-file` | `driveId`, `itemId` | Delete file/folder |
| `search-files` | `query` | Search files |

### To Do Tasks
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-todo-task-lists` | | List task lists |
| `list-todo-tasks` | `todoTaskListId` | List tasks in list |
| `get-todo-task` | `todoTaskListId`, `taskId` | Get task details |
| `create-todo-task` | `todoTaskListId`, `body` | Create task |
| `update-todo-task` | `todoTaskListId`, `taskId`, `body` | Update task |
| `complete-todo-task` | `todoTaskListId`, `taskId` | Mark complete |
| `delete-todo-task` | `todoTaskListId`, `taskId` | Delete task |

### Contacts
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-outlook-contacts` | `top` | List contacts |
| `get-outlook-contact` | `contactId` | Get contact details |
| `search-people` | `search` | Search contacts |
| `create-outlook-contact` | `body` (contact object) | Create contact |

### OneNote
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-notebooks` | | List notebooks |
| `list-notebook-sections` | `notebookId` | List sections |
| `list-section-pages` | `notebookId`, `sectionId` | List pages |
| `get-page-content` | `pageId` | Get page content |

### Organization Features (--org-mode)
| Tool | Parameters | Description |
|------|------------|-------------|
| `list-teams` | | List Teams |
| `list-team-channels` | `teamId` | List channels |
| `send-channel-message` | `teamId`, `channelId`, `message` | Post to channel |
| `list-chats` | | List chats |
| `send-chat-message` | `chatId`, `message` | Send chat message |
| `list-sites` | | List SharePoint sites |
| `list-site-documents` | `siteId` | List site documents |

## Data Models

### Email Message
```json
{
  "id": "string",
  "subject": "string",
  "body": {
    "contentType": "Text|HTML",
    "content": "string"
  },
  "from": {"emailAddress": {"address": "string", "name": "string"}},
  "toRecipients": [{"emailAddress": {"address": "string", "name": "string"}}],
  "ccRecipients": [],
  "bccRecipients": [],
  "hasAttachments": boolean,
  "importance": "Low|Normal|High",
  "isRead": boolean,
  "createdDateTime": "ISO 8601 datetime",
  "lastModifiedDateTime": "ISO 8601 datetime"
}
```

### Calendar Event
```json
{
  "id": "string",
  "subject": "string",
  "body": {
    "contentType": "Text|HTML",
    "content": "string"
  },
  "start": {
    "dateTime": "ISO 8601 datetime",
    "timeZone": "string"
  },
  "end": {
    "dateTime": "ISO 8601 datetime", 
    "timeZone": "string"
  },
  "location": {"displayName": "string"},
  "isAllDay": boolean,
  "isOrganizer": boolean,
  "attendees": [],
  "organizer": {"emailAddress": {"address": "string", "name": "string"}},
  "createdDateTime": "ISO 8601 datetime",
  "lastModifiedDateTime": "ISO 8601 datetime"
}
```

### Task
```json
{
  "id": "string",
  "title": "string",
  "body": {
    "contentType": "Text|HTML",
    "content": "string"
  },
  "isCompleted": boolean,
  "dueDateTime": {
    "dateTime": "ISO 8601 datetime",
    "timeZone": "string"
  },
  "importance": "Low|Normal|High",
  "createdDateTime": "ISO 8601 datetime",
  "lastModifiedDateTime": "ISO 8601 datetime"
}
```

### Contact
```json
{
  "id": "string",
  "displayName": "string",
  "givenName": "string",
  "surname": "string",
  "emailAddresses": [{"address": "string", "name": "string"}],
  "phones": [{"type": "Mobile|Business|Home", "number": "string"}],
  "jobTitle": "string",
  "companyName": "string",
  "department": "string"
}
```

## Error Handling

Common error responses:
```json
{
  "error": {
    "code": "string",
    "message": "string",
    "details": []
  }
}
```

### Common Error Codes
- `AADSTS700016`: Application not found
- `AADSTS7000218`: Invalid client secret  
- `AADSTS70005`: Invalid scope
- `ErrorItemNotFound`: Resource not found
- `ErrorPermissionDenied`: Insufficient permissions

## Rate Limits

Microsoft Graph API has rate limits:
- 60 requests per minute per application
- 10,000 requests per 10 minutes per application
- 60 requests per minute per user

## Best Practices

1. **Batch Operations**: Use batch requests when possible
2. **Pagination**: Handle `@odata.nextLink` for large result sets
3. **Select Fields**: Use `$select` to minimize data transfer
4. **Error Handling**: Implement retry logic for transient errors
5. **Caching**: Cache frequently accessed data where appropriate
6. **Security**: Never expose client secrets in client-side code

## Configuration Options

### Environment Variables
| Variable | Description | Default |
|----------|-------------|---------|
| `MS365_MCP_CLIENT_ID` | Azure AD app client ID | Required for headless |
| `MS365_MCP_CLIENT_SECRET` | Azure AD app secret | Required for headless |
| `MS365_MCP_TENANT_ID` | Tenant ID | `common` |
| `MS365_MCP_ORG_MODE` | Enable org features | `false` |
| `MS365_MCP_READ_ONLY` | Read-only mode | `false` |
| `MS365_MCP_OUTPUT_FORMAT` | Output format | `json` |
| `MS365_MCP_PORT` | HTTP server port | `3365` |

### Server Arguments
| Argument | Description |
|----------|-------------|
| `--org-mode` | Enable organization features |
| `--read-only` | Disable write operations |
| `--toon` | Use token-efficient output |
| `--discovery` | Start with minimal tools |
| `--http <port>` | Run in HTTP mode |
| `--auth-only` | Device code auth only |