# MS 365 Skill API Documentation

## Overview

This document provides API documentation for the MS 365 Skill, including the Python CLI wrapper and MCP server integration.

## Python CLI API

### Core Functions

#### `call_mcp(method: str, params: dict = None) -> dict`

Call the MCP server via stdio using JSON-RPC protocol.

**Parameters:**
- `method` (str): The MCP tool name to call
- `params` (dict, optional): Tool arguments

**Returns:**
- `dict`: Response data or error information

**Example:**
```python
result = call_mcp("list-mail-messages", {"top": 10})
```

#### `format_output(data: dict, compact: bool = False)`

Format and display output data.

**Parameters:**
- `data` (dict): Data to format
- `compact` (bool): Use compact JSON formatting

### Command Functions

#### Authentication Commands

##### `cmd_login(args)`
- **Purpose**: Login via device code flow
- **Usage**: `python3 ms365_cli.py login`
- **Notes**: Interactive authentication process

##### `cmd_status(args)`
- **Purpose**: Check authentication status
- **Usage**: `python3 ms365_cli.py status`
- **Returns**: Authentication status and account information

##### `cmd_accounts(args)`
- **Purpose**: List cached accounts
- **Usage**: `python3 ms365_cli.py accounts`
- **Returns**: List of cached authentication accounts

##### `cmd_user(args)`
- **Purpose**: Get current user information
- **Usage**: `python3 ms365_cli.py user`
- **Returns**: User profile data from Microsoft Graph

#### Email Commands

##### `cmd_mail_list(args)`
- **Purpose**: List emails in a folder
- **Usage**: `python3 ms365_cli.py mail list [--top N] [--folder FOLDER_ID]`
- **Parameters:**
  - `--top` (int): Maximum number of emails to return
  - `--folder` (str): Folder ID (optional, defaults to inbox)
- **Returns**: List of email messages

##### `cmd_mail_read(args)`
- **Purpose**: Read a specific email
- **Usage**: `python3 ms365_cli.py mail read MESSAGE_ID`
- **Parameters:**
  - `MESSAGE_ID` (str): Unique identifier for the email
- **Returns**: Full email content and metadata

##### `cmd_mail_send(args)`
- **Purpose**: Send a new email
- **Usage**: `python3 ms365_cli.py mail send --to RECIPIENT --subject SUBJECT --body BODY`
- **Parameters:**
  - `--to` (str): Recipient email address (required)
  - `--subject` (str): Email subject (required)
  - `--body` (str): Email body content (required)
- **Returns**: Send operation result

#### Calendar Commands

##### `cmd_calendar_list(args)`
- **Purpose**: List calendar events
- **Usage**: `python3 ms365_cli.py calendar list [--top N]`
- **Parameters:**
  - `--top` (int): Maximum number of events to return
- **Returns**: List of calendar events

##### `cmd_calendar_create(args)`
- **Purpose**: Create a new calendar event
- **Usage**: `python3 ms365_cli.py calendar create --subject SUBJECT --start START --end END [--body BODY] [--timezone TIMEZONE]`
- **Parameters:**
  - `--subject` (str): Event title (required)
  - `--start` (str): Start time in ISO 8601 format (required)
  - `--end` (str): End time in ISO 8601 format (required)
  - `--body` (str): Event description (optional)
  - `--timezone` (str): Timezone (optional, defaults to America/Chicago)
- **Returns**: Created event details

#### File Commands

##### `cmd_files_list(args)`
- **Purpose**: List OneDrive files in a folder
- **Usage**: `python3 ms365_cli.py files list [--path PATH]`
- **Parameters:**
  - `--path` (str): Folder path (optional, defaults to root)
- **Returns**: List of files and folders

#### Task Commands

##### `cmd_tasks_list(args)`
- **Purpose**: List To Do task lists
- **Usage**: `python3 ms365_cli.py tasks lists`
- **Returns**: List of task lists

##### `cmd_tasks_get(args)`
- **Purpose**: Get tasks from a specific list
- **Usage**: `python3 ms365_cli.py tasks get LIST_ID`
- **Parameters:**
  - `LIST_ID` (str): Task list identifier
- **Returns**: List of tasks in the specified list

##### `cmd_tasks_create(args)`
- **Purpose**: Create a new task
- **Usage**: `python3 ms365_cli.py tasks create LIST_ID --title TITLE [--due DUE_DATE]`
- **Parameters:**
  - `LIST_ID` (str): Task list identifier
  - `--title` (str): Task title (required)
  - `--due` (str): Due date in ISO format (optional)
- **Returns**: Created task details

#### Contact Commands

##### `cmd_contacts_list(args)`
- **Purpose**: List contacts
- **Usage**: `python3 ms365_cli.py contacts list [--top N]`
- **Parameters:**
  - `--top` (int): Maximum number of contacts to return
- **Returns**: List of contacts

##### `cmd_contacts_search(args)`
- **Purpose**: Search contacts
- **Usage**: `python3 ms365_cli.py contacts SEARCH_QUERY`
- **Parameters:**
  - `SEARCH_QUERY` (str): Search string
- **Returns**: Matching contacts

## MCP Server Integration

### Available Tools

The skill integrates with the `@softeria/ms-365-mcp-server` which provides the following tools:

#### Email Tools
- `list-mail-messages`: List emails
- `get-mail-message`: Get specific email
- `send-mail`: Send email
- `reply-mail`: Reply to email
- `forward-mail`: Forward email
- `delete-mail`: Delete email
- `move-mail`: Move email to folder
- `search-mail`: Search emails

#### Calendar Tools
- `list-calendar-events`: List events
- `get-calendar-event`: Get specific event
- `create-calendar-event`: Create event
- `update-calendar-event`: Update event
- `delete-calendar-event`: Delete event
- `list-calendars`: List calendars

#### File Tools
- `list-folder-files`: List files in folder
- `get-file`: Get file metadata
- `download-file`: Download file content
- `upload-file`: Upload file
- `create-folder`: Create folder
- `delete-file`: Delete file/folder
- `search-files`: Search files

#### Task Tools
- `list-todo-task-lists`: List task lists
- `list-todo-tasks`: List tasks in list
- `create-todo-task`: Create task
- `update-todo-task`: Update task
- `complete-todo-task`: Complete task
- `delete-todo-task`: Delete task

#### Contact Tools
- `list-outlook-contacts`: List contacts
- `get-outlook-contact`: Get specific contact
- `search-people`: Search contacts
- `create-outlook-contact`: Create contact

#### Note Tools
- `list-notebooks`: List OneNote notebooks
- `list-notebook-sections`: List notebook sections
- `list-section-pages`: List section pages
- `get-page-content`: Get page content

#### Organization Tools (with --org-mode)
- `list-teams`: List Microsoft Teams
- `list-team-channels`: List team channels
- `send-channel-message`: Send message to channel
- `list-chats`: List chat conversations
- `send-chat-message`: Send chat message
- `list-sites`: List SharePoint sites
- `list-site-files`: List site documents

## Environment Variables

### Required for Azure AD Authentication
- `MS365_MCP_CLIENT_ID`: Azure AD application client ID
- `MS365_MCP_CLIENT_SECRET`: Azure AD application client secret
- `MS365_MCP_TENANT_ID`: Azure AD tenant ID (use "consumers" for personal accounts)

### Optional Configuration
- `MS365_MCP_ORG_MODE`: Enable organization features (Teams, SharePoint)
- `MS365_MCP_READ_ONLY`: Restrict to read-only operations
- `MS365_MCP_OUTPUT_FORMAT`: Set output format (e.g., "toon")
- `MS365_MCP_PORT`: HTTP server port (default: 3365)

## Error Handling

The CLI includes comprehensive error handling:

1. **Authentication Errors**: Clear messages about login issues
2. **Network Errors**: Timeout and connection error handling
3. **API Errors**: Graceful handling of Microsoft Graph API errors
4. **JSON Parsing**: Fallback to raw text if JSON parsing fails

## Return Format

All commands return data in the following format:

```json
{
  "success": true,
  "data": {...},
  "error": null
}
```

Or for errors:

```json
{
  "success": false,
  "data": null,
  "error": {
    "message": "Error description",
    "code": "ERROR_CODE",
    "details": {...}
  }
}
```