---
name: ms365
description: |
  Microsoft 365 integration via Graph API. Access Outlook email, Calendar,
  OneDrive, To Do, Planner, Contacts, OneNote, and more.

  Personal accounts: Email, Calendar, OneDrive, To Do, Contacts, OneNote
  Work accounts (--org-mode): All above plus Teams, SharePoint, Shared Mailboxes

  Uses ms-365-mcp-server MCP server via mcporter.
homepage: https://github.com/Softeria/ms-365-mcp-server
metadata:
  clawdbot:
    emoji: "📧"
    primaryEnv: "MS365_MCP_CLIENT_ID"
    requires:
      bins: ["node", "npx"]
    install:
      - id: ms365-mcp
        kind: node
        package: "@softeria/ms-365-mcp-server"
        label: "Install ms-365-mcp-server"
---

# Microsoft 365 Integration

Access and manage your Microsoft 365 services through natural conversation.

## Capabilities

### Email (Outlook)
- List, read, search emails
- Send emails with attachments
- Delete, move, mark as read/unread
- Search across mailboxes

### Calendar
- View upcoming events
- Create, update, delete events
- Check availability
- Manage recurring events

### OneDrive
- List files and folders
- Upload and download files
- Create folders, delete items
- Share files

### To Do
- List tasks and lists
- Create, update, complete tasks
- Set due dates and reminders

### Contacts
- List and search contacts
- View contact details
- Create new contacts

### OneNote
- List notebooks and sections
- Read page content
- Create notes

### Work/Organization Features (requires --org-mode)
- Teams messaging
- SharePoint site access
- Shared mailboxes

## Usage Examples

### Reading Email

"Check my recent emails"
"Show me unread emails from today"
"Search for emails from john@example.com"
"Find emails with attachments about the quarterly report"

### Sending Email

"Send an email to jane@example.com with subject 'Meeting Tomorrow' and body 'Can we meet at 2pm?'"
"Reply to that email saying I'll be there"

### Calendar

"What's on my calendar today?"
"Show me my meetings for this week"
"Create a meeting called 'Team Standup' tomorrow at 9am for 30 minutes"
"Cancel my 3pm meeting"

### OneDrive

"List files in my Documents folder"
"Download the budget spreadsheet"
"Upload this file to OneDrive"

### To Do

"Show my tasks"
"Add a task: Review quarterly report by Friday"
"Mark the 'Send invoice' task as complete"

## Technical Details

This skill uses `mcporter` to communicate with the `ms-365-mcp-server` MCP server.

### Direct mcporter Commands

```bash
# List available tools
mcporter list ms365

# Email operations
mcporter call ms365.list_messages folder=inbox limit=10
mcporter call ms365.get_message id="MESSAGE_ID"
mcporter call ms365.send_message to="user@example.com" subject="Hello" body="Message"
mcporter call ms365.search_messages query="from:boss subject:urgent"

# Calendar operations
mcporter call ms365.list_events
mcporter call ms365.create_event subject="Meeting" start="2026-01-15T10:00:00" end="2026-01-15T11:00:00"
mcporter call ms365.update_event id="EVENT_ID" subject="Updated Meeting"
mcporter call ms365.delete_event id="EVENT_ID"

# OneDrive operations
mcporter call ms365.list_files path="/"
mcporter call ms365.list_files path="/Documents"
mcporter call ms365.download_file path="/Documents/report.pdf"
mcporter call ms365.upload_file path="/Documents/new-file.txt" content="File content"

# To Do operations
mcporter call ms365.list_task_lists
mcporter call ms365.list_tasks list_id="LIST_ID"
mcporter call ms365.create_task list_id="LIST_ID" title="New task" due_date="2026-01-20"
mcporter call ms365.complete_task list_id="LIST_ID" task_id="TASK_ID"

# Contacts
mcporter call ms365.list_contacts limit=20
mcporter call ms365.search_contacts query="John"
```

## Setup

### Prerequisites

1. Node.js 20+ installed
2. Microsoft account (personal or work)
3. Azure AD app registration (for headless operation)

### Authentication Options

**Option 1: Azure AD App (Recommended for Docker/headless)**

Set these environment variables:
- `MS365_MCP_CLIENT_ID` - Azure app client ID
- `MS365_MCP_CLIENT_SECRET` - Azure app client secret
- `MS365_MCP_TENANT_ID` - `consumers` for personal, or your tenant ID

**Option 2: Device Code Flow (Interactive)**

No setup needed. On first use, follow the prompts to authenticate via browser.

See README.md for detailed Azure AD setup instructions.

## Environment Variables

| Variable | Description |
|----------|-------------|
| `MS365_MCP_CLIENT_ID` | Azure AD application client ID |
| `MS365_MCP_CLIENT_SECRET` | Azure AD application client secret |
| `MS365_MCP_TENANT_ID` | Tenant ID (`consumers` for personal accounts) |
| `MS365_MCP_ORG_MODE` | Set to `true` for Teams/SharePoint access |
| `MS365_MCP_OUTPUT_FORMAT` | Set to `toon` for 30-60% token reduction |

## Troubleshooting

### "Authentication required"
Run device code auth or verify Azure AD credentials are set correctly.

### "Insufficient permissions"
Check Azure AD app has required Graph API permissions.

### "Token expired"
Re-authenticate using device code flow or refresh credentials.
