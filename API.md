# MS 365 Skill API Documentation

This document provides detailed API documentation for the MS 365 Skill components.

## Overview

The MS 365 Skill provides two main interfaces:
1. **Python CLI Wrapper** (`ms365_cli.py`) - Command-line interface for direct interaction
2. **MCP Server Integration** - Integration with Clawdbot via mcporter

## Python CLI API

### Core Functions

#### `call_mcp(method: str, params: dict = None) -> dict`

Call the MCP server via stdio using JSON-RPC protocol.

**Parameters:**
- `method` (str): The MCP tool method name to call
- `params` (dict, optional): Method parameters. Defaults to None.

**Returns:**
- `dict`: Response from the MCP server, parsed as JSON or raw text

**Raises:**
- `subprocess.TimeoutExpired`: If the MCP server doesn't respond within 60 seconds
- `Exception`: For other subprocess or JSON parsing errors

**Example:**
```python
result = call_mcp("list_messages", {"top": 5})
print(result)
# {'value': [{'id': 'msg1', 'subject': 'Hello'}]}
```

#### `format_output(data: dict, compact: bool = False)`

Format and display MCP server response output.

**Parameters:**
- `data` (dict): The response data from MCP server
- `compact` (bool, optional): Use compact JSON formatting. Defaults to False.

**Example:**
```python
format_output({"value": [{"id": "msg1"}]}, compact=True)
# {"value": [{"id": "msg1"}]}
```

### Command Functions

#### Authentication Commands

##### `cmd_login(args)`
Initiate device code flow authentication.

**Parameters:**
- `args`: argparse namespace (unused)

**Example:**
```bash
python3 ms365_cli.py login
```

##### `cmd_status(args)`
Check current authentication status.

**Parameters:**
- `args`: argparse namespace (unused)

**Returns:**
- Authentication status and account information

**Example:**
```bash
python3 ms365_cli.py status
# {"authenticated": true, "account": "user@example.com"}
```

##### `cmd_accounts(args)`
List all cached authentication accounts.

**Parameters:**
- `args`: argparse namespace (unused)

**Returns:**
- List of cached accounts

**Example:**
```bash
python3 ms365_cli.py accounts
# {"accounts": [{"id": "acc1", "email": "user@example.com"}]}
```

##### `cmd_user(args)`
Get current authenticated user information.

**Parameters:**
- `args`: argparse namespace (unused)

**Returns:**
- User profile information

**Example:**
```bash
python3 ms365_cli.py user
# {"user": {"displayName": "John Doe", "mail": "john@example.com"}}
```

#### Email Commands

##### `cmd_mail_list(args)`
List emails in the user's mailbox.

**Parameters:**
- `args.top` (int, optional): Maximum number of emails to return
- `args.folder` (str, optional): Folder ID to list emails from

**Returns:**
- List of email messages

**Example:**
```bash
python3 ms365_cli.py mail list --top 10
```

##### `cmd_mail_read(args)`
Read the full content of a specific email.

**Parameters:**
- `args.id` (str): The message ID of the email to read

**Returns:**
- Complete email details

**Example:**
```bash
python3 ms365_cli.py mail read AAMkAG...=
```

##### `cmd_mail_send(args)`
Send a new email message.

**Parameters:**
- `args.to` (str): Recipient email address
- `args.subject` (str): Email subject line
- `args.body` (str): Email body content

**Returns:**
- Message ID and confirmation

**Example:**
```bash
python3 ms365_cli.py mail send --to "john@example.com" --subject "Hello" --body "How are you?"
```

#### Calendar Commands

##### `cmd_calendar_list(args)`
List calendar events.

**Parameters:**
- `args.top` (int, optional): Maximum number of events to return

**Returns:**
- List of calendar events

**Example:**
```bash
python3 ms365_cli.py calendar list --top 5
```

##### `cmd_calendar_create(args)`
Create a new calendar event.

**Parameters:**
- `args.subject` (str): Event subject/title
- `args.start` (str): Start time in ISO 8601 format
- `args.end` (str): End time in ISO 8601 format
- `args.body` (str, optional): Event description
- `args.timezone` (str, optional): Timezone (default: America/Chicago)

**Returns:**
- Event ID and confirmation

**Example:**
```bash
python3 ms365_cli.py calendar create --subject "Team Meeting" --start "2026-01-15T10:00:00" --end "2026-01-15T11:00:00"
```

#### File Commands

##### `cmd_files_list(args)`
List files and folders in OneDrive.

**Parameters:**
- `args.path` (str, optional): Folder path to list (default: root)

**Returns:**
- List of files and folders

**Example:**
```bash
python3 ms365_cli.py files list --path "Documents"
```

#### Task Commands

##### `cmd_tasks_list(args)`
List all To Do task lists.

**Parameters:**
- `args`: argparse namespace (unused)

**Returns:**
- List of task lists

**Example:**
```bash
python3 ms365_cli.py tasks lists
```

##### `cmd_tasks_get(args)`
Get tasks from a specific task list.

**Parameters:**
- `args.list_id` (str): The ID of the task list

**Returns:**
- List of tasks in the specified list

**Example:**
```bash
python3 ms365_cli.py tasks get "list1"
```

##### `cmd_tasks_create(args)`
Create a new task in a task list.

**Parameters:**
- `args.list_id` (str): The ID of the task list
- `args.title` (str): Task title/description
- `args.due` (str, optional): Due date in ISO 8601 format

**Returns:**
- Task ID and confirmation

**Example:**
```bash
python3 ms365_cli.py tasks create "list1" --title "Review budget" --due "2026-01-20"
```

#### Contact Commands

##### `cmd_contacts_list(args)`
List contacts from Outlook.

**Parameters:**
- `args.top` (int, optional): Maximum number of contacts to return

**Returns:**
- List of contacts

**Example:**
```bash
python3 ms365_cli.py contacts list --top 10
```

##### `cmd_contacts_search(args)`
Search for contacts by name or email.

**Parameters:**
- `args.query` (str): Search query string

**Returns:**
- List of matching contacts

**Example:**
```bash
python3 ms365_cli.py contacts search "John"
```

## MCP Server Integration

The skill integrates with the `@softeria/ms-365-mcp-server` MCP server through mcporter. Available tools include:

### Email Tools
- `list-mail-messages` - List emails
- `get-mail-message` - Get email details
- `send-mail` - Send email
- `reply-message` - Reply to email
- `forward-message` - Forward email
- `delete-message` - Delete email
- `move-message` - Move email to folder
- `search-messages` - Search emails

### Calendar Tools
- `list-calendar-events` - List events
- `get-calendar-event` - Get event details
- `create-calendar-event` - Create event
- `update-calendar-event` - Update event
- `delete-calendar-event` - Delete event
- `list-calendars` - List calendars

### OneDrive Tools
- `list-folder-files` - List files/folders
- `get-file` - Get file metadata
- `download-file` - Download file
- `upload-file` - Upload file
- `create-folder` - Create folder
- `delete-file` - Delete file/folder
- `search-files` - Search files

### To Do Tools
- `list-todo-task-lists` - List task lists
- `list-todo-tasks` - List tasks
- `create-todo-task` - Create task
- `update-todo-task` - Update task
- `complete-todo-task` - Complete task
- `delete-todo-task` - Delete task

### Contact Tools
- `list-outlook-contacts` - List contacts
- `get-outlook-contact` - Get contact details
- `search-people` - Search contacts
- `create-outlook-contact` - Create contact

### OneNote Tools
- `list-notebooks` - List notebooks
- `list-sections` - List sections
- `list-pages` - List pages
- `get-page-content` - Get page content

### Organization Tools (with --org-mode)
- `list-teams` - List Teams
- `list-channels` - List channels
- `send-channel-message` - Post to channel
- `list-chats` - List chats
- `send-chat-message` - Send chat message
- `list-sites` - List SharePoint sites
- `list-site-files` - List site documents

## Environment Variables

### Required for Azure AD Authentication
- `MS365_MCP_CLIENT_ID` - Azure AD app client ID
- `MS365_MCP_CLIENT_SECRET` - Azure AD app client secret
- `MS365_MCP_TENANT_ID` - Tenant ID (use "consumers" for personal accounts)

### Optional Configuration
- `MS365_MCP_ORG_MODE` - Enable organization features (Teams, SharePoint)
- `MS365_MCP_READ_ONLY` - Disable write operations
- `MS365_MCP_OUTPUT_FORMAT` - Output format (toon for token efficiency)
- `MS365_MCP_PORT` - HTTP server port (default: 3365)

## Error Handling

The CLI wrapper provides consistent error handling:

1. **Timeout Errors**: Requests timeout after 60 seconds
2. **Authentication Errors**: Clear error messages for auth failures
3. **API Errors**: Microsoft Graph API errors are returned as-is
4. **Network Errors**: Connection and network issues are reported

## Return Format

All functions return a dictionary with the following structure:

```python
{
    "success": bool,           # Operation success status
    "data": dict,             # Response data (if successful)
    "error": str,             # Error message (if failed)
    "raw": str                # Raw response (for debugging)
}
```

## Examples

### Basic Usage
```python
from ms365_cli import call_mcp, format_output

# List emails
result = call_mcp("list-mail-messages", {"top": 5})
format_output(result)

# Create calendar event
event_data = {
    "subject": "Team Meeting",
    "start": {"dateTime": "2026-01-15T10:00:00", "timeZone": "America/Chicago"},
    "end": {"dateTime": "2026-01-15T11:00:00", "timeZone": "America/Chicago"}
}
result = call_mcp("create-calendar-event", {"body": event_data})
format_output(result)
```

### Error Handling
```python
result = call_mcp("list-mail-messages", {"top": 5})
if "error" in result:
    print(f"Error: {result['error']}")
else:
    format_output(result)
```