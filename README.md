# MS 365 Skill for Clawdbot

> **Your Microsoft 365, accessible through conversation.**

A [Clawdbot](https://github.com/clawdbot/clawdbot) skill that brings Microsoft 365 services—Outlook, Calendar, OneDrive, To Do, and more—to your AI assistant via the Microsoft Graph API.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)
- [Advanced Configuration](#advanced-configuration)
- [Tool Reference](#tool-reference)
- [Support & License](#support--license)

---

## Features

| Service | Capabilities |
|---------|-------------|
| **Email** | Read, compose, search, reply, forward, and organize messages |
| **Calendar** | View schedule, create meetings, manage events |
| **OneDrive** | Browse, upload, download, and manage files |
| **To Do** | Create tasks, set due dates, organize lists |
| **Contacts** | Search and view your address book |
| **OneNote** | Access notebooks and pages |
| **Teams** *(org mode)* | Send messages, access chats and channels |
| **SharePoint** *(org mode)* | Browse sites and documents |

---

## Installation

### 1. Install the Skill

Clone this repository into your Clawdbot skills directory:

```bash
# For workspace skills
cd <your-workspace>/skills
git clone https://github.com/cvsloane/m365-skill ms365

# Or for global installation
cd ~/.clawdbot/skills
git clone https://github.com/cvsloane/m365-skill ms365
```

### 2. Install the MCP Server

The skill requires the Microsoft 365 MCP server:

```bash
npm install -g @softeria/ms-365-mcp-server
```

> **Note:** The included Python CLI wrapper works with Python 3.6+ using only standard library modules—no additional Python dependencies required.

### 3. Configure Authentication

Choose the authentication method that fits your use case:

#### Option A: Azure AD App Registration (Recommended)

**Best for:** Headless deployments, automation, Docker containers, and production use.

This method provides persistent authentication without requiring interactive login each time.

<details>
<summary><b>Step 1: Create an Azure AD App</b></summary>

1. Visit the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** → **App registrations**
3. Click **New registration**
4. Configure the application:
   - **Name**: `Clawdbot MS365` (or your preferred name)
   - **Supported account types**:
     - Select "Personal Microsoft accounts only" for personal Outlook/Hotmail
     - Select "Accounts in any organizational directory and personal" for work + personal
   - **Redirect URI**: Select "Web" and enter `http://localhost:3365/callback`
5. Click **Register**

</details>

<details>
<summary><b>Step 2: Grant API Permissions</b></summary>

1. In your app registration, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:

| Permission | Purpose |
|------------|---------|
| `Mail.ReadWrite` | Read, send, and manage emails |
| `Mail.Send` | Send emails |
| `Calendars.ReadWrite` | Create and manage calendar events |
| `Files.ReadWrite` | Access OneDrive files |
| `Tasks.ReadWrite` | Manage To Do tasks |
| `Contacts.Read` | Access contacts |
| `Notes.Read` | Read OneNote notebooks |
| `User.Read` | Access basic profile info |

**For organization features** (Teams, SharePoint), also add:

| Permission | Purpose |
|------------|---------|
| `Chat.ReadWrite` | Teams messaging |
| `Sites.Read.All` | SharePoint access |
| `ChannelMessage.Send` | Post to Teams channels |

4. Click **Grant admin consent** if you have administrator privileges

</details>

<details>
<summary><b>Step 3: Create a Client Secret</b></summary>

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: `Clawdbot`
4. Choose expiration (24 months recommended)
5. Click **Add**
6. **⚠️ IMPORTANT**: Copy the secret value immediately—it is shown only once!

</details>

<details>
<summary><b>Step 4: Collect Your Credentials</b></summary>

From the **Overview** page, note:

| Credential | Environment Variable |
|------------|---------------------|
| **Application (client) ID** | `MS365_MCP_CLIENT_ID` |
| **Directory (tenant) ID** | `MS365_MCP_TENANT_ID` |
| **Client secret** (from Step 3) | `MS365_MCP_CLIENT_SECRET` |

> **Tip:** For personal accounts only, use `consumers` as the tenant ID.

</details>

<details>
<summary><b>Step 5: Set Environment Variables</b></summary>

Add these to your environment (`.env` file, Docker environment, or deployment platform):

```bash
MS365_MCP_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
MS365_MCP_CLIENT_SECRET=your-secret-value
MS365_MCP_TENANT_ID=consumers
```

</details>

#### Option B: Device Code Flow

**Best for:** Quick testing, personal use, or when you can't create an Azure app.

No Azure setup required—just run any MS365 command and follow the prompts:

```bash
# This will display a device code
python3 ms365_cli.py login
```

1. Open the provided URL (usually `https://microsoft.com/devicelogin`)
2. Enter the code shown in your terminal
3. Sign in with your Microsoft account
4. Tokens are automatically cached for future use

> **Note:** Device code flow requires interactive access and won't work in pure headless environments like Docker containers without a terminal.

### 4. Configure mcporter

Add the MS365 server to your mcporter configuration. Create or edit `~/.clawdbot/mcporter.json`:

**Using stdio mode (recommended):**
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

**Using HTTP mode:**
```json
{
  "servers": {
    "ms365": {
      "url": "http://localhost:3365"
    }
  }
}
```

### 5. Verify Your Installation

Test the connection to Microsoft 365:

```bash
# Check server accessibility
mcporter list ms365

# Test email access
mcporter call ms365.list_messages limit=5

# Test calendar access
mcporter call ms365.list_events
```

---

## Usage

### Natural Language with Clawdbot

Once configured, simply talk to Clawdbot naturally:

- *"Check my email"*
- *"What meetings do I have tomorrow?"*
- *"Send an email to john@example.com about the project update"*
- *"Add a task to review the budget by Friday"*
- *"Show my OneDrive files"*
- *"Find emails from Sarah last week"*

### Direct mcporter Commands

For precise control or scripting:

```bash
# List recent emails
mcporter call ms365.list_messages limit=10

# List today's calendar events
mcporter call ms365.list_events top=5

# Search for emails
mcporter call ms365.search_messages query="from:boss@company.com"

# List OneDrive files
mcporter call ms365.list_files
```

### Python CLI Wrapper

The included `ms365_cli.py` provides a convenient command-line interface:

```bash
# Check your email
python3 ms365_cli.py mail list --top 10

# View calendar events
python3 ms365_cli.py calendar list

# Send an email
python3 ms365_cli.py mail send \
  --to "user@example.com" \
  --subject "Project Update" \
  --body "Here's the latest status..."

# Create a task
python3 ms365_cli.py tasks create \
  LIST_ID \
  --title "Review quarterly report" \
  --due "2026-03-31"
```

See [SKILL.md](./SKILL.md) for complete CLI documentation.

---

## Troubleshooting

### Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| **AADSTS700016: Application not found** | Invalid client ID | Verify `MS365_MCP_CLIENT_ID` matches your Azure app registration |
| **AADSTS7000218: Invalid client secret** | Expired or incorrect secret | Create a new secret in Azure portal and update `MS365_MCP_CLIENT_SECRET` |
| **Insufficient privileges** | Missing API permissions | Add required permissions in Azure portal and grant admin consent if needed |
| **Token expired / 401 Unauthorized** | Expired authentication | Re-authenticate using device code flow or refresh Azure app credentials |
| **mcporter can't find ms365 server** | Configuration issue | Check `mcporter.json` syntax; verify npm package is installed with `npm list -g @softeria/ms-365-mcp-server` |

### Still Having Issues?

1. **Test the server directly:**
   ```bash
   npx -y @softeria/ms-365-mcp-server
   ```
   Should start without errors.

2. **Check your environment variables:**
   ```bash
   echo $MS365_MCP_CLIENT_ID
   echo $MS365_MCP_TENANT_ID
   ```

3. **Clear token cache and re-authenticate:**
   ```bash
   rm -rf ~/.clawdbot/ms365-cache
   python3 ms365_cli.py login
   ```

---

## Advanced Configuration

### Organization/Work Account Features

Enable Teams, SharePoint, and shared mailbox access:

1. **Set the environment variable:**
   ```bash
   export MS365_MCP_ORG_MODE=true
   ```

2. **Update mcporter configuration:**
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

3. **Ensure your Azure app has organization permissions** (see installation Step 2)

### Token-Efficient Output (TOON)

Reduce token usage by 30-60% for large responses:

```bash
export MS365_MCP_OUTPUT_FORMAT=toon
```

### Read-Only Mode

Prevent all write operations for enhanced security:

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

Start with minimal tools and expand as needed (useful for initial exploration):

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

## Tool Reference

### Email (Outlook)

| Tool | Description |
|------|-------------|
| `list_messages` | List emails in a folder (default: Inbox) |
| `get_message` | Retrieve full email content by ID |
| `send_message` | Compose and send a new email |
| `reply_message` | Reply to an existing email |
| `forward_message` | Forward an email to another recipient |
| `delete_message` | Move email to trash |
| `move_message` | Move email between folders |
| `search_messages` | Search emails by query |

### Calendar

| Tool | Description |
|------|-------------|
| `list_events` | List upcoming calendar events |
| `get_event` | Get detailed event information |
| `create_event` | Create a new calendar event |
| `update_event` | Modify an existing event |
| `delete_event` | Remove a calendar event |
| `list_calendars` | List all available calendars |

### OneDrive

| Tool | Description |
|------|-------------|
| `list_files` | List files in a folder |
| `get_file` | Get file metadata |
| `download_file` | Download file content |
| `upload_file` | Upload a new file |
| `create_folder` | Create a new folder |
| `delete_file` | Delete a file or folder |
| `search_files` | Search for files |

### To Do

| Tool | Description |
|------|-------------|
| `list_task_lists` | Get all task lists |
| `list_tasks` | List tasks in a specific list |
| `create_task` | Create a new task |
| `update_task` | Modify a task |
| `complete_task` | Mark a task as complete |
| `delete_task` | Remove a task |

### Contacts

| Tool | Description |
|------|-------------|
| `list_contacts` | List contacts |
| `get_contact` | Get contact details |
| `search_contacts` | Search contacts |
| `create_contact` | Create a new contact |

### OneNote

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all notebooks |
| `list_sections` | List sections in a notebook |
| `list_pages` | List pages in a section |
| `get_page_content` | Retrieve page content |

### Organization Tools (`--org-mode`)

| Tool | Description |
|------|-------------|
| `list_teams` | List Microsoft Teams |
| `list_channels` | List channels in a team |
| `send_channel_message` | Post a message to a channel |
| `list_chats` | List chat conversations |
| `send_chat_message` | Send a direct message |
| `list_sites` | List SharePoint sites |
| `list_site_files` | List documents in a site |

---

## Support & License

### Getting Help

- **MCP Server Issues:** [github.com/Softeria/ms-365-mcp-server/issues](https://github.com/Softeria/ms-365-mcp-server/issues)
- **Clawdbot Issues:** [github.com/clawdbot/clawdbot/issues](https://github.com/clawdbot/clawdbot/issues)
- **Skill Issues:** Open an issue on this repository

### License

MIT License — feel free to use, modify, and share!

### Development Status

See [project_status.md](./project_status.md) for recent development activity.
