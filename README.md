# Microsoft 365 Skill for Clawdbot

> Seamlessly integrate Microsoft 365 services—Outlook, Calendar, OneDrive, Teams, and more—into your Clawdbot workflows via the Microsoft Graph API.

---

## Table of Contents

- [Features](#features)
- [Quick Start](#quick-start)
- [Installation](#installation)
  - [1. Copy Skill to Clawdbot](#1-copy-skill-to-clawdbot)
  - [2. Install Dependencies](#2-install-dependencies)
  - [3. Set Up Authentication](#3-set-up-authentication)
  - [4. Configure mcporter](#4-configure-mcporter)
  - [5. Verify Installation](#5-verify-installation)
- [Usage](#usage)
  - [Natural Language with Clawdbot](#natural-language-with-clawdbot)
  - [Direct mcporter Commands](#direct-mcporter-commands)
  - [Python CLI Wrapper](#python-cli-wrapper)
- [Organization Features](#organization-features)
- [Advanced Configuration](#advanced-configuration)
- [Troubleshooting](#troubleshooting)
- [Available Tools Reference](#available-tools-reference)
- [Support & Resources](#support--resources)

---

## Features

| Service | Capabilities |
|---------|-------------|
| **Email (Outlook)** | Read, compose, search, organize, and manage messages |
| **Calendar** | View schedules, create meetings, set reminders |
| **OneDrive** | Browse folders, upload documents, download files |
| **To Do** | Create lists, manage tasks, track deadlines |
| **Contacts** | Search address book, view contact details |
| **OneNote** | Access notebooks, read pages and sections |
| **Teams** *(org mode)* | Send messages, participate in chats and channels |
| **SharePoint** *(org mode)* | Access sites, manage shared documents |

---

## Quick Start

For the impatient—get running in 5 minutes:

```bash
# 1. Clone the skill
cd ~/.clawdbot/skills && git clone https://github.com/cvsloane/m365-skill ms365

# 2. Install the MCP server
npm install -g @softeria/ms-365-mcp-server

# 3. Authenticate (interactive device code flow)
python3 ms365/ms365_cli.py login

# 4. Test it
mcporter call ms365.list_messages limit=5
```

That's it! No Azure setup required for personal use. For headless deployments or organization accounts, continue to the full installation guide below.

---

## Installation

### 1. Copy Skill to Clawdbot

Choose the appropriate location based on your setup:

```bash
# For workspace-specific skills
cd <your-workspace>/skills
git clone https://github.com/cvsloane/m365-skill ms365

# For global/user-wide skills
cd ~/.clawdbot/skills
git clone https://github.com/cvsloane/m365-skill ms365
```

### 2. Install Dependencies

The skill requires the Microsoft 365 MCP server from npm:

```bash
npm install -g @softeria/ms-365-mcp-server
```

**Requirements:**
- Node.js 16+ (for the MCP server)
- Python 3.6+ (standard library only—no pip dependencies needed)

### 3. Set Up Authentication

Choose the method that fits your deployment scenario:

#### Option A: Azure AD App Registration (Recommended for Production)

This method is **required** for Docker containers, headless servers, or any automated deployment where interactive login isn't possible.

**Step 1: Create the App Registration**

1. Navigate to the [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure as follows:
   - **Name**: `Clawdbot MS365` (or your preferred name)
   - **Supported account types**:
     - *Personal Microsoft accounts only* → For personal Outlook/Live accounts
     - *Accounts in any organizational directory and personal* → For work + personal
   - **Redirect URI**: Select **Web** and enter `http://localhost:3365/callback`
5. Click **Register**

**Step 2: Configure API Permissions**

In your app registration, go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:

**Core Permissions (Required for Personal Use):**

| Permission | Purpose |
|------------|---------|
| `Mail.ReadWrite` | Read and send emails |
| `Mail.Send` | Send emails (separate permission) |
| `Calendars.ReadWrite` | Manage calendar events |
| `Files.ReadWrite` | Access OneDrive files |
| `Tasks.ReadWrite` | Manage To Do tasks |
| `Contacts.Read` | Read contacts |
| `Notes.Read` | Access OneNote notebooks |
| `User.Read` | Basic profile information |

**Organization Permissions (Enable with `--org-mode`):**

| Permission | Purpose |
|------------|---------|
| `Chat.ReadWrite` | Teams direct messages |
| `ChannelMessage.Send` | Post to Teams channels |
| `Sites.Read.All` | SharePoint site access |

After adding permissions, click **Grant admin consent** if you have administrative privileges.

**Step 3: Create a Client Secret**

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter a description (e.g., `Clawdbot`)
4. Choose expiration (24 months recommended)
5. Click **Add**
6. **⚠️ Important**: Copy the secret value **immediately**—it's shown only once!

**Step 4: Collect Your Credentials**

From the **Overview** page, note:
- **Application (client) ID** → Set as `MS365_MCP_CLIENT_ID`
- **Directory (tenant) ID** → Set as `MS365_MCP_TENANT_ID`
  - Use `consumers` for personal accounts only
- **Client secret value** (from Step 3) → Set as `MS365_MCP_CLIENT_SECRET`

**Step 5: Configure Environment Variables**

Add to your environment (`.env` file, Docker, or deployment platform):

```bash
export MS365_MCP_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export MS365_MCP_CLIENT_SECRET="your-secret-value-here"
export MS365_MCP_TENANT_ID="consumers"  # or your tenant ID
```

#### Option B: Device Code Flow (Quick Testing)

For interactive testing without Azure setup:

1. Run any MS365 command (e.g., `python3 ms365_cli.py mail list`)
2. You'll see a prompt: *"To sign in, use a web browser to open https://microsoft.com/devicelogin"*
3. Enter the code displayed
4. Sign in with your Microsoft account
5. Tokens are cached locally for future use

> **Note:** Device code flow requires interactive access and won't work in pure headless environments like Docker containers.

### 4. Configure mcporter

Add the MS365 server to your mcporter configuration. Edit `~/.clawdbot/mcporter.json`:

**Stdio Mode (Recommended):**

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

**HTTP Mode (Alternative):**

```json
{
  "servers": {
    "ms365": {
      "url": "http://localhost:3365"
    }
  }
}
```

> **Pro tip:** Use HTTP mode when you want to run the MCP server separately or share it across multiple clients.

### 5. Verify Installation

Confirm everything is working:

```bash
# Check server connectivity
mcporter list ms365

# Test email access
mcporter call ms365.list_messages limit=5

# Test calendar access
mcporter call ms365.list_events top=5
```

If these commands return data, you're ready to go!

---

## Usage

### Natural Language with Clawdbot

Once configured, interact with Microsoft 365 naturally through Clawdbot:

| You say | Clawdbot does |
|---------|---------------|
| "Check my email" | Lists recent messages |
| "What meetings do I have tomorrow?" | Shows tomorrow's calendar |
| "Send an email to john@example.com about the project update" | Composes and sends the message |
| "Add a task to review the budget by Friday" | Creates a To Do item |
| "Show my OneDrive files" | Lists your files |
| "Find emails from Sarah last week" | Searches with filters |

### Direct mcporter Commands

For scripting or precise control, use mcporter directly:

```bash
# List 10 most recent emails
mcporter call ms365.list_messages limit=10

# Get calendar for next week
mcporter call ms365.list_events top=20

# Search emails
mcporter call ms365.search_messages query="from:boss@company.com"

# Upload a file
mcporter call ms365.upload_file path="/path/to/file.txt"
```

### Python CLI Wrapper

The included `ms365_cli.py` provides a familiar command-line interface:

```bash
# Email operations
python3 ms365_cli.py mail list --top 10
python3 ms365_cli.py mail send --to "user@example.com" --subject "Hello" --body "Message"

# Calendar operations
python3 ms365_cli.py calendar list
python3 ms365_cli.py calendar create --subject "Team Standup" --start "2026-02-15T09:00:00" --end "2026-02-15T09:30:00"

# OneDrive operations
python3 ms365_cli.py files list --path "Documents"

# Task management
python3 ms365_cli.py tasks lists
python3 ms365_cli.py tasks create "LIST_ID" --title "Review quarterly report" --due "2026-02-20"
```

> **Note:** When running from within Clawdbot, use the full path to the script (e.g., `/root/clawd/skills/ms365/ms365_cli.py`).

---

## Organization Features

Unlock Teams, SharePoint, and enhanced collaboration features for work accounts:

### 1. Enable Organization Mode

Set the environment variable:

```bash
export MS365_MCP_ORG_MODE=true
```

### 2. Update mcporter Configuration

```json
{
  "servers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--org-mode"]
    }
  }
}
```

### 3. Verify Azure Permissions

Ensure your Azure app registration has the organization permissions listed in the authentication section, and that your IT admin has granted consent.

---

## Advanced Configuration

### Token-Efficient Output (TOON Format)

Reduce API token usage by 30-60% with structured output:

```bash
export MS365_MCP_OUTPUT_FORMAT=toon
```

Ideal for high-volume automation or when working with large datasets.

### Read-Only Mode

Prevent accidental modifications by disabling all write operations:

```json
{
  "servers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--read-only"]
    }
  }
}
```

### Discovery Mode

Start with minimal tools and expand capabilities on demand—useful for constrained environments:

```json
{
  "servers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--discovery"]
    }
  }
}
```

---

## Troubleshooting

### Common Errors

#### "AADSTS700016: Application not found"

**Cause:** The `MS365_MCP_CLIENT_ID` doesn't match an existing app registration.

**Solution:**
- Verify the client ID in the [Azure Portal](https://portal.azure.com) → App registrations
- Ensure you're looking at the correct Azure AD tenant
- Check for typos or extra spaces in the environment variable

#### "AADSTS7000218: Invalid client secret"

**Cause:** The client secret has expired or was entered incorrectly.

**Solution:**
1. Navigate to your app registration → Certificates & secrets
2. Check if the secret has expired
3. Create a new secret if needed
4. Update `MS365_MCP_CLIENT_SECRET` with the new value

#### "Insufficient privileges" / "403 Forbidden"

**Cause:** Missing API permissions or admin consent not granted.

**Solution:**
- Add required permissions in Azure Portal → API permissions
- Click **Grant admin consent** (if you're an admin)
- For work accounts, ask your IT admin to approve the app

#### "Token expired" / "401 Unauthorized"

**Cause:** Authentication token has expired or been invalidated.

**Solution:**
- Re-authenticate using device code flow: `scripts/auth-device.sh`
- Clear the token cache and re-authenticate
- Check if the client secret has expired (for Azure AD app method)
- Verify system clock is accurate (tokens are time-sensitive)

#### "mcporter can't find ms365 server"

**Cause:** Configuration issue or missing dependency.

**Solution:**
```bash
# Verify mcporter.json syntax
python3 -m json.tool ~/.clawdbot/mcporter.json

# Check npm package is installed
npm list -g @softeria/ms-365-mcp-server

# Test server directly
npx -y @softeria/ms-365-mcp-server --help
```

### Debug Mode

Enable verbose logging to diagnose issues:

```bash
# Run server directly with debug output
DEBUG=* npx -y @softeria/ms-365-mcp-server

# Check authentication status
python3 ms365_cli.py status
```

---

## Available Tools Reference

### Email

| Tool | Description |
|------|-------------|
| `list_messages` | List emails in a folder |
| `get_message` | Retrieve full email content |
| `send_message` | Send a new email |
| `reply_message` | Reply to an email |
| `forward_message` | Forward an email |
| `delete_message` | Move to deleted items |
| `move_message` | Move to another folder |
| `search_messages` | Search with filters |

### Calendar

| Tool | Description |
|------|-------------|
| `list_events` | List upcoming events |
| `get_event` | Get event details |
| `create_event` | Schedule a new event |
| `update_event` | Modify existing event |
| `delete_event` | Cancel an event |
| `list_calendars` | List all calendars |

### OneDrive

| Tool | Description |
|------|-------------|
| `list_files` | List files in a folder |
| `get_file` | Get file metadata |
| `download_file` | Download file content |
| `upload_file` | Upload a file |
| `create_folder` | Create a new folder |
| `delete_file` | Delete file or folder |
| `search_files` | Search for files |

### To Do

| Tool | Description |
|------|-------------|
| `list_task_lists` | List all task lists |
| `list_tasks` | Get tasks from a list |
| `create_task` | Add a new task |
| `update_task` | Modify task details |
| `complete_task` | Mark as completed |
| `delete_task` | Remove a task |

### Contacts

| Tool | Description |
|------|-------------|
| `list_contacts` | List contacts |
| `get_contact` | Get contact details |
| `search_contacts` | Search contacts |
| `create_contact` | Add new contact |

### OneNote

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all notebooks |
| `list_sections` | List notebook sections |
| `list_pages` | List section pages |
| `get_page_content` | Read page content |

### Organization (--org-mode)

| Tool | Description |
|------|-------------|
| `list_teams` | List Microsoft Teams |
| `list_channels` | List team channels |
| `send_channel_message` | Post to a channel |
| `list_chats` | List chat conversations |
| `send_chat_message` | Send direct message |
| `list_sites` | List SharePoint sites |
| `list_site_files` | List site documents |

---

## Support & Resources

### Getting Help

- **MCP Server Issues:** [Softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server/issues)
- **Clawdbot Issues:** [clawdbot/clawdbot](https://github.com/clawdbot/clawdbot/issues)
- **Skill Issues:** [cvsloane/m365-skill](https://github.com/cvsloane/m365-skill/issues)

### Useful Links

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure Portal](https://portal.azure.com)
- [npm Package: @softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server)

### Development

For information about contributing or modifying this skill, see:
- [CONTRIBUTING.md](./CONTRIBUTING.md) - Contribution guidelines
- [CLAUDE.md](./CLAUDE.md) - Developer notes for Claude Code
- [project_status.md](./project_status.md) - Recent development activity

---

## License

MIT License — feel free to use, modify, and share.

---

<p align="center">
  <em>Built with ❤️ for the Clawdbot community</em>
</p>
