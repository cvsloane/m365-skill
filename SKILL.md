# Microsoft 365 Integration

Access Microsoft 365 services—Email, Calendar, OneDrive, To Do tasks, and Contacts—via the Microsoft Graph API.

---

## Description

This skill enables natural language interaction with Microsoft 365 services through Clawdbot. Whether managing emails, scheduling meetings, organizing tasks, or accessing files, users can accomplish their work using conversational commands.

---

## Activation

This skill activates when users mention:

**Primary keywords:** outlook, email, calendar, onedrive, microsoft, office 365, o365, ms365, to do, microsoft tasks

**Natural phrases:** my meetings, my emails, schedule meeting, send email, check calendar

---

## Configuration

### Authentication Methods

**For Interactive Use (Device Code Flow):**

No configuration required. Authentication is cached after the first login. Simply run any command and follow the device code prompts.

**For Headless/Automated Operation:**

Set these environment variables with your Azure AD app credentials:

| Variable | Description | Example |
|----------|-------------|---------|
| `MS365_MCP_CLIENT_ID` | Azure AD application ID | `a1b2c3d4-e5f6-...` |
| `MS365_MCP_CLIENT_SECRET` | Client secret from Azure | `your-secret-here` |
| `MS365_MCP_TENANT_ID` | Tenant identifier | `consumers` (personal) or tenant GUID |

> **Note:** See [README.md](./README.md) for detailed Azure AD setup instructions.

---

## Available Commands

All commands use the `ms365_cli.py` script. When running from Clawdbot, use the full path (e.g., `python3 /root/clawd/skills/ms365/ms365_cli.py`).

### Authentication & Account

```bash
# Start device code login flow
python3 ms365_cli.py login

# Check authentication status
python3 ms365_cli.py status

# List cached Microsoft accounts
python3 ms365_cli.py accounts

# Get current user profile
python3 ms365_cli.py user
```

### Email (Outlook)

```bash
# List recent emails (default: 10)
python3 ms365_cli.py mail list [--top N]

# Read a specific email by ID
python3 ms365_cli.py mail read <MESSAGE_ID>

# Send a new email
python3 ms365_cli.py mail send \
  --to "recipient@example.com" \
  --subject "Subject Line" \
  --body "Message content"
```

**Examples:**

| User Request | Command Executed |
|-------------|------------------|
| "Check my email" | `mail list --top 10` |
| "Show me my last 5 emails" | `mail list --top 5` |
| "Read the email with ID abc123" | `mail read abc123` |
| "Email john@example.com about the project" | `mail send --to john@example.com --subject "Project Update" --body "..."` |

### Calendar

```bash
# List upcoming events (default: 10)
python3 ms365_cli.py calendar list [--top N]

# Create a new calendar event
python3 ms365_cli.py calendar create \
  --subject "Meeting Title" \
  --start "2026-02-15T10:00:00" \
  --end "2026-02-15T11:00:00" \
  [--body "Meeting description"] \
  [--timezone "America/Chicago"]
```

**Time Format:** ISO 8601 datetime (`YYYY-MM-DDTHH:MM:SS`). Default timezone is America/Chicago.

**Examples:**

| User Request | Command Executed |
|-------------|------------------|
| "What meetings do I have today?" | `calendar list --top 20` |
| "Schedule a 1-hour meeting tomorrow at 2pm" | `calendar create --subject "Meeting" --start "2026-02-16T14:00:00" --end "2026-02-16T15:00:00"` |

### OneDrive Files

```bash
# List files in root directory
python3 ms365_cli.py files list

# List files in a specific folder
python3 ms365_cli.py files list --path "Documents/Projects"
```

**Examples:**

| User Request | Command Executed |
|-------------|------------------|
| "Show my OneDrive files" | `files list` |
| "List files in my Documents folder" | `files list --path "Documents"` |

### To Do Tasks

```bash
# List all task lists
python3 ms365_cli.py tasks lists

# Get tasks from a specific list
python3 ms365_cli.py tasks get <LIST_ID>

# Create a new task
python3 ms365_cli.py tasks create <LIST_ID> \
  --title "Task description" \
  [--due "2026-02-20"]
```

**Examples:**

| User Request | Command Executed |
|-------------|------------------|
| "Show my task lists" | `tasks lists` |
| "Add a task to review the budget" | `tasks lists` → `tasks create <id> --title "Review budget"` |
| "Add a task due Friday" | `tasks create <id> --title "Task" --due "2026-02-20"` |

### Contacts

```bash
# List contacts (default: 20)
python3 ms365_cli.py contacts list [--top N]

# Search contacts by name or email
python3 ms365_cli.py contacts search "John Smith"
```

**Examples:**

| User Request | Command Executed |
|-------------|------------------|
| "Find John's email" | `contacts search "John"` |
| "List my contacts" | `contacts list --top 50` |

---

## Usage Examples

### Scenario 1: Morning Email Check

**User:** "Check my outlook email"

**Action:** Run `mail list --top 10` and summarize the results.

### Scenario 2: Meeting Discovery

**User:** "What meetings do I have today?"

**Action:** Run `calendar list` and present events chronologically.

### Scenario 3: Email Composition

**User:** "Send an email to john@company.com about the project update"

**Action:** 
1. Confirm recipient and intent
2. Run `mail send --to john@company.com --subject "Project Update" --body "..."`
3. Confirm successful delivery

### Scenario 4: File Access

**User:** "Show my OneDrive files"

**Action:** Run `files list` and display the directory structure.

### Scenario 5: Task Management

**User:** "Add a task to review the budget"

**Action:**
1. Run `tasks lists` to get available lists
2. Ask user which list (or use default)
3. Run `tasks create <list_id> --title "Review the budget"`

---

## Prompts & Best Practices

When helping users with Microsoft 365 operations:

### Always Do

- **Use the CLI wrapper** (`ms365_cli.py`) for all direct operations
- **Check authentication first** if a command fails with auth errors
- **Guide through login** if the user isn't authenticated (use device code flow)
- **Use ISO 8601 format** for all datetimes (`YYYY-MM-DDTHH:MM:SS`)
- **Confirm before sending**—always verify recipient and content for emails
- **Present task lists first** before creating tasks so users can choose

### Never Do

- Don't assume folder names or list IDs—always fetch them first
- Don't send emails without explicit confirmation
- Don't use ambiguous time formats (always include timezone context)

### Pro Tips

- The default timezone is America/Chicago—ask if the user needs a different one
- For calendar events, include timezone in the output to avoid confusion
- When listing emails, highlight unread messages or those from important senders
- Task lists often have IDs that aren't human-readable—present names, not IDs

---

## Attribution

This skill is built on the excellent **ms-365-mcp-server** by Softeria.

| Resource | Link |
|----------|------|
| **NPM Package** | [@softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server) |
| **GitHub Repository** | [Softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server) |
| **License** | MIT |

---

## Additional Resources

- [Full Setup Guide (README.md)](./README.md)
- [Contributing Guidelines](./CONTRIBUTING.md)
- [Developer Notes (CLAUDE.md)](./CLAUDE.md)
