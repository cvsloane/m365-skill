# Microsoft 365 Integration

## Description

Access your Microsoft 365 world—Outlook email, Calendar, OneDrive files, To Do tasks, Contacts, and more—through the Microsoft Graph API. This skill enables natural conversation with your Microsoft 365 data.

## Activation

This skill activates when users mention:

- `outlook`, `email`, `mail`
- `calendar`, `schedule`, `meetings`, `appointment`
- `onedrive`, `my files`, `cloud storage`
- `microsoft`, `office 365`, `o365`, `ms365`
- `to do`, `tasks`, `todo list`
- `contacts`, `address book`

**Example triggers:**
- *"Check my email"*
- *"What meetings do I have today?"*
- *"Show my OneDrive files"*
- *"Add a task to review the budget"*

## Configuration

### Authentication

**For interactive use:** No configuration required. The skill uses device code flow and caches authentication after first login.

**For headless/automated operation:** Set these environment variables:

| Variable | Required | Description |
|----------|----------|-------------|
| `MS365_MCP_CLIENT_ID` | Yes | Azure AD application client ID |
| `MS365_MCP_CLIENT_SECRET` | Yes | Azure AD application secret |
| `MS365_MCP_TENANT_ID` | Yes | Tenant ID (use `"consumers"` for personal accounts) |

### Optional Settings

| Variable | Description |
|----------|-------------|
| `MS365_MCP_ORG_MODE=true` | Enable Teams and SharePoint features |
| `MS365_MCP_OUTPUT_FORMAT=toon` | Use token-efficient output format |

---

## Available Commands

All commands use the `ms365_cli.py` script. When running from Clawdbot, use the full path:

```bash
# Example paths (adjust to your installation)
/root/clawd/skills/ms365/ms365_cli.py        # Container
~/workspace/skills/ms365/ms365_cli.py        # Local workspace
~/.clawdbot/skills/ms365/ms365_cli.py        # Global install
```

### Authentication Commands

#### Login

Start device code authentication flow:

```bash
python3 ms365_cli.py login
```

**What happens:**
1. A device code and URL are displayed
2. Open the URL in your browser
3. Enter the device code
4. Sign in with your Microsoft account
5. Authentication is cached for future use

#### Check Status

Verify authentication status:

```bash
python3 ms365_cli.py status
```

#### List Cached Accounts

View all authenticated accounts:

```bash
python3 ms365_cli.py accounts
```

#### Get Current User

Display profile information for the authenticated user:

```bash
python3 ms365_cli.py user
```

---

### Email (Outlook)

#### List Recent Emails

```bash
python3 ms365_cli.py mail list [--top N]
```

**Options:**
- `--top N` — Number of emails to retrieve (default: 10)

**Example:**
```bash
python3 ms365_cli.py mail list --top 20
```

#### Read a Specific Email

```bash
python3 ms365_cli.py mail read MESSAGE_ID
```

**How to find the message ID:** Use `mail list` and copy the `id` field from the email you want to read.

#### Send an Email

```bash
python3 ms365_cli.py mail send \
  --to "recipient@example.com" \
  --subject "Your Subject" \
  --body "Your message body"
```

**Options:**
- `--to` — Recipient email address (required)
- `--subject` — Email subject line (required)
- `--body` — Message body content (required)

**Tip:** For multi-line bodies, use a text file:
```bash
python3 ms365_cli.py mail send \
  --to "team@company.com" \
  --subject "Weekly Update" \
  --body "$(cat update.txt)"
```

---

### Calendar

#### List Upcoming Events

```bash
python3 ms365_cli.py calendar list [--top N]
```

**Options:**
- `--top N` — Number of events to retrieve (default: 10)

#### Create a New Event

```bash
python3 ms365_cli.py calendar create \
  --subject "Meeting Name" \
  --start "2026-01-15T10:00:00" \
  --end "2026-01-15T11:00:00" \
  [--body "Meeting description"] \
  [--timezone "America/Chicago"]
```

**Required:**
- `--subject` — Event title
- `--start` — Start time in ISO 8601 format
- `--end` — End time in ISO 8601 format

**Optional:**
- `--body` — Event description
- `--timezone` — Timezone (default: America/Chicago)

**Common Timezones:**
- `America/New_York` — Eastern Time
- `America/Chicago` — Central Time
- `America/Denver` — Mountain Time
- `America/Los_Angeles` — Pacific Time
- `UTC` — Coordinated Universal Time
- `Europe/London` — GMT/BST

**Example:**
```bash
python3 ms365_cli.py calendar create \
  --subject "Project Review" \
  --start "2026-02-15T14:00:00" \
  --end "2026-02-15T15:00:00" \
  --body "Quarterly project status review" \
  --timezone "America/New_York"
```

---

### OneDrive Files

#### List Files

```bash
# Root folder
python3 ms365_cli.py files list

# Specific folder
python3 ms365_cli.py files list --path "Documents/Projects"
```

**Tips:**
- Paths are relative to your OneDrive root
- Use forward slashes `/` regardless of operating system
- Folder names are case-sensitive

---

### To Do Tasks

#### List Task Lists

Get all your task lists (required before adding tasks):

```bash
python3 ms365_cli.py tasks lists
```

**Sample output:**
```json
{
  "value": [
    { "id": "list-123", "displayName": "Tasks" },
    { "id": "list-456", "displayName": "Work" },
    { "id": "list-789", "displayName": "Shopping" }
  ]
}
```

#### List Tasks in a List

```bash
python3 ms365_cli.py tasks get LIST_ID
```

#### Create a New Task

```bash
python3 ms365_cli.py tasks create LIST_ID \
  --title "Task title" \
  [--due "2026-01-20"]
```

**Example workflow:**
```bash
# 1. Get list ID
python3 ms365_cli.py tasks lists

# 2. Create task (use the actual ID from step 1)
python3 ms365_cli.py tasks create list-123 \
  --title "Review quarterly report" \
  --due "2026-03-31"
```

---

### Contacts

#### List Contacts

```bash
python3 ms365_cli.py contacts list [--top N]
```

#### Search Contacts

```bash
python3 ms365_cli.py contacts search "search term"
```

**Example:**
```bash
# Search for John
python3 ms365_cli.py contacts search "John"

# Search by email domain
python3 ms365_cli.py contacts search "@company.com"
```

---

## Usage Examples

### Scenario: Checking Morning Email

**User:** *"Check my outlook email"*

**Agent:**
```bash
python3 ms365_cli.py mail list --top 10
```

### Scenario: Finding Today's Meetings

**User:** *"What meetings do I have today?"*

**Agent:**
```bash
python3 ms365_cli.py calendar list
```

### Scenario: Sending a Project Update

**User:** *"Send an email to john@company.com about the project update"*

**Agent:**
```bash
python3 ms365_cli.py mail send \
  --to "john@company.com" \
  --subject "Project Update" \
  --body "Hi John,

Here's the latest update on the project...

Best regards"
```

### Scenario: Browsing Files

**User:** *"Show my OneDrive files"*

**Agent:**
```bash
python3 ms365_cli.py files list
```

### Scenario: Creating a Task

**User:** *"Add a task to review the budget"*

**Agent:**
```bash
# First, get available task lists
python3 ms365_cli.py tasks lists

# Then create the task in the appropriate list
python3 ms365_cli.py tasks create LIST_ID \
  --title "Review the budget"
```

---

## Prompts

When helping users with Microsoft 365:

### Authentication
- **Check status first** if commands fail with authentication errors
- If not authenticated, guide users through the device code login flow
- Remind users that authentication persists after first login

### Email
- Always confirm recipient addresses before sending
- Confirm email content (subject and body) for important messages
- Use `--top` to limit results for better readability

### Calendar
- Use **ISO 8601 datetime format**: `YYYY-MM-DDTHH:MM:SS`
- Default timezone is `America/Chicago`—confirm if user is in a different zone
- Validate datetime formats before creating events

### Tasks
- **Always list task lists first**—users need to choose which list to use
- Suggest due dates if users don't specify them
- Remind users that task lists are personal—choose the right one

### Files
- Paths in OneDrive use forward slashes `/`
- Folder names are case-sensitive
- Remind users that files are stored in their Microsoft cloud

---

## Attribution

This skill is powered by **ms-365-mcp-server** from Softeria.

| Resource | Link |
|----------|------|
| **NPM Package** | [@softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server) |
| **GitHub Repository** | [github.com/Softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server) |
| **License** | MIT |

---

## Quick Reference Card

```bash
# AUTHENTICATION
python3 ms365_cli.py login                    # Device code login
python3 ms365_cli.py status                   # Check auth status
python3 ms365_cli.py user                     # Get user info

# EMAIL
python3 ms365_cli.py mail list --top 10       # List emails
python3 ms365_cli.py mail read MESSAGE_ID     # Read email
python3 ms365_cli.py mail send --to "..." --subject "..." --body "..."

# CALENDAR
python3 ms365_cli.py calendar list            # List events
python3 ms365_cli.py calendar create --subject "..." --start "2026-01-15T10:00:00" --end "..."

# FILES
python3 ms365_cli.py files list               # List files
python3 ms365_cli.py files list --path "..."  # List folder

# TASKS
python3 ms365_cli.py tasks lists              # Get task lists
python3 ms365_cli.py tasks get LIST_ID        # List tasks
python3 ms365_cli.py tasks create LIST_ID --title "..." [--due "..."]

# CONTACTS
python3 ms365_cli.py contacts list            # List contacts
python3 ms365_cli.py contacts search "..."    # Search contacts
```
