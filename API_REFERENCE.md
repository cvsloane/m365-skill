# MS 365 Skill API Reference

This document provides detailed API documentation for the MS 365 Skill components.

## Python CLI Interface (`ms365_cli.py`)

### Core Functions

#### `call_mcp(method: str, params: dict = None) -> dict`

Call the MCP server via stdio using JSON-RPC protocol.

**Parameters:**
- `method` (str): The MCP tool name to call
- `params` (dict, optional): Tool parameters. Defaults to None.

**Returns:**
- `dict`: Response data from the MCP server or error information

**Example:**
```python
result = call_mcp("list_messages", {"top": 10})
print(result)
```

#### `format_output(data: dict, compact: bool = False)`

Format and display output data.

**Parameters:**
- `data` (dict): Data to format and display
- `compact` (bool): Whether to use compact JSON formatting. Defaults to False.

### Command Functions

#### Authentication Commands

##### `cmd_login(args)`
Login via device code flow.

**Parameters:**
- `args`: Parsed command line arguments

##### `cmd_status(args)`
Check authentication status.

**Parameters:**
- `args`: Parsed command line arguments

**Returns:**
- Authentication status information from the MCP server

##### `cmd_accounts(args)`
List cached accounts.

**Parameters:**
- `args`: Parsed command line arguments

**Returns:**
- List of cached authentication accounts

##### `cmd_user(args)`
Get current user information.

**Parameters:**
- `args`: Parsed command line arguments

**Returns:**
- Current user profile information

#### Email Commands

##### `cmd_mail_list(args)`
List emails in a folder.

**Parameters:**
- `args.top` (int, optional): Maximum number of emails to return. Defaults to 10.
- `args.folder` (str, optional): Folder ID. Defaults to inbox.

**Returns:**
- List of email messages

##### `cmd_mail_read(args)`
Read a specific email.

**Parameters:**
- `args.id` (str): Message ID of the email to read

**Returns:**
- Full email content and metadata

##### `cmd_mail_send(args)`
Send an email.

**Parameters:**
- `args.to` (str): Recipient email address
- `args.subject` (str): Email subject
- `args.body` (str): Email body content

**Returns:**
- Send operation result

#### Calendar Commands

##### `cmd_calendar_list(args)`
List calendar events.

**Parameters:**
- `args.top` (int, optional): Maximum number of events to return. Defaults to 10.

**Returns:**
- List of calendar events

##### `cmd_calendar_create(args)`
Create a calendar event.

**Parameters:**
- `args.subject` (str): Event subject
- `args.start` (str): Start time in ISO 8601 format
- `args.end` (str): End time in ISO 8601 format
- `args.body` (str, optional): Event description
- `args.timezone` (str, optional): Timezone. Defaults to "America/Chicago"

**Returns:**
- Created event information

#### File Commands

##### `cmd_files_list(args)`
List OneDrive files in a folder.

**Parameters:**
- `args.path` (str, optional): Folder path. Defaults to "root".

**Returns:**
- List of files and folders

#### Task Commands

##### `cmd_tasks_list(args)`
List To Do task lists.

**Parameters:**
- `args`: Parsed command line arguments

**Returns:**
- List of task lists

##### `cmd_tasks_get(args)`
Get tasks from a specific list.

**Parameters:**
- `args.list_id` (str): Task list ID

**Returns:**
- List of tasks in the specified list

##### `cmd_tasks_create(args)`
Create a new task.

**Parameters:**
- `args.list_id` (str): Task list ID
- `args.title` (str): Task title
- `args.due` (str, optional): Due date in ISO format

**Returns:**
- Created task information

#### Contact Commands

##### `cmd_contacts_list(args)`
List contacts.

**Parameters:**
- `args.top` (int, optional): Maximum number of contacts to return. Defaults to 20.

**Returns:**
- List of contacts

##### `cmd_contacts_search(args)`
Search contacts.

**Parameters:**
- `args.query` (str): Search query string

**Returns:**
- Search results matching the query

## MCP Server Interface

The skill communicates with the Microsoft 365 Graph API through the `@softeria/ms-365-mcp-server` package.

### Available Tools

#### Email Tools
- `list-mail-messages`: List emails with optional filtering
- `get-mail-message`: Retrieve specific email content
- `send-mail`: Send new emails
- `reply-mail`: Reply to existing emails
- `forward-mail`: Forward emails
- `delete-mail-message`: Delete emails
- `move-mail-message`: Move emails between folders
- `search-mail-messages`: Search emails

#### Calendar Tools
- `list-calendar-events`: List calendar events
- `get-calendar-event`: Get event details
- `create-calendar-event`: Create new events
- `update-calendar-event`: Update existing events
- `delete-calendar-event`: Delete events
- `list-calendars`: List available calendars

#### OneDrive Tools
- `list-folder-files`: List files and folders
- `get-file`: Get file metadata
- `download-file`: Download file content
- `upload-file`: Upload files
- `create-folder`: Create new folders
- `delete-file`: Delete files and folders
- `search-files`: Search for files

#### To Do Tools
- `list-todo-task-lists`: List task lists
- `list-todo-tasks`: List tasks in a list
- `create-todo-task`: Create new tasks
- `update-todo-task`: Update existing tasks
- `complete-todo-task`: Mark tasks as complete
- `delete-todo-task`: Delete tasks

#### Contact Tools
- `list-outlook-contacts`: List contacts
- `get-outlook-contact`: Get contact details
- `search-people`: Search contacts
- `create-outlook-contact`: Create new contacts

#### OneNote Tools
- `list-notebooks`: List notebooks
- `list-notebook-sections`: List sections in a notebook
- `list-section-pages`: List pages in a section
- `get-page-content`: Get page content

#### Organization Tools (with --org-mode)
- `list-teams`: List Microsoft Teams
- `list-team-channels`: List channels in a team
- `send-channel-message`: Post messages to channels
- `list-chats`: List chat conversations
- `send-chat-message`: Send chat messages
- `list-sites`: List SharePoint sites
- `list-site-documents`: List documents in a site

## Environment Variables

### Required for Azure AD Authentication
- `MS365_MCP_CLIENT_ID`: Azure AD application client ID
- `MS365_MCP_CLIENT_SECRET`: Azure AD application client secret
- `MS365_MCP_TENANT_ID`: Azure AD tenant ID (use "consumers" for personal accounts)

### Optional Configuration
- `MS365_MCP_ORG_MODE`: Enable organization features (Teams, SharePoint)
- `MS365_MCP_READ_ONLY`: Disable write operations
- `MS365_MCP_OUTPUT_FORMAT`: Set output format (e.g., "toon")
- `MS365_MCP_PORT`: HTTP server port (default: 3365)

## Error Handling

The CLI functions handle various error conditions:

1. **Authentication Errors**: Returns error information with guidance
2. **Network Errors**: Handles timeouts and connection issues
3. **API Errors**: Returns Microsoft Graph API error responses
4. **Parameter Errors**: Validates input parameters before making API calls

## Return Format

All functions return a dictionary with the following structure:

```python
{
    "success": bool,
    "data": any,           # Response data or None
    "error": str,          # Error message if failed
    "raw": str            # Raw response for debugging
}
```