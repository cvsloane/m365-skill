# Documentation Verification Report

**Date**: 2026-01-31
**Verifier**: Claude (Sonnet)
**Repository**: m365-skill

## Summary

Found **significant discrepancies** between documentation and actual MCP server tool names. The documentation uses simplified/snake_case tool names that don't match the actual kebab-case tool names used by the @softeria/ms-365-mcp-server.

## Critical Issues Found

### 1. Tool Name Mismatches (CRITICAL)

The README's "Available Tools Reference" section uses incorrect tool names. The actual server uses kebab-case (e.g., `list-mail-messages`), not snake_case (e.g., `list_messages`).

**Email Tools - Documented vs Actual:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_messages` | `list-mail-messages` | ❌ MISMATCH |
| `get_message` | `get-mail-message` | ❌ MISMATCH |
| `send_message` | `send-mail` | ❌ MISMATCH |
| `reply_message` | `create-draft-email`? | ❌ WRONG/UNCLEAR |
| `forward_message` | ? | ❌ NOT FOUND |
| `delete_message` | `delete-mail-message` | ❌ MISMATCH |
| `move_message` | `move-mail-message` | ❌ MISMATCH |
| `search_messages` | `search-query` | ❌ MISMATCH |

**Calendar Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_events` | `list-calendar-events` | ❌ MISMATCH |
| `get_event` | `get-calendar-event` | ❌ MISMATCH |
| `create_event` | `create-calendar-event` | ❌ MISMATCH |
| `update_event` | `update-calendar-event` | ❌ MISMATCH |
| `delete_event` | `delete-calendar-event` | ❌ MISMATCH |
| `list_calendars` | `list-calendars` | ❌ MISMATCH |

**OneDrive Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_files` | `list-folder-files` | ❌ MISMATCH |
| `get_file` | ? | ❌ UNCLEAR |
| `download_file` | `download-onedrive-file-content` | ❌ MISMATCH |
| `upload_file` | `upload-file-content` or `upload-new-file` | ❌ MISMATCH |
| `create_folder` | ? | ❌ NOT FOUND |
| `delete_file` | `delete-onedrive-file` | ❌ MISMATCH |
| `search_files` | `search-query` | ❌ MISMATCH |

**To Do Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_task_lists` | `list-todo-task-lists` | ❌ MISMATCH |
| `list_tasks` | `list-todo-tasks` | ❌ MISMATCH |
| `create_task` | `create-todo-task` | ❌ MISMATCH |
| `update_task` | `update-todo-task` | ❌ MISMATCH |
| `complete_task` | ? | ❌ NOT FOUND |
| `delete_task` | `delete-todo-task` | ❌ MISMATCH |

**Contact Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_contacts` | `list-outlook-contacts` | ❌ MISMATCH |
| `get_contact` | `get-outlook-contact` | ❌ MISMATCH |
| `search_contacts` | `search-query` | ❌ MISMATCH |
| `create_contact` | `create-outlook-contact` | ❌ MISMATCH |

**OneNote Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_notebooks` | `list-onenote-notebooks` | ❌ MISMATCH |
| `list_sections` | `list-onenote-notebook-sections` | ❌ MISMATCH |
| `list_pages` | `list-onenote-section-pages` | ❌ MISMATCH |
| `get_page_content` | `get-onenote-page-content` | ❌ MISMATCH |

**Organization Tools:**
| Documented | Actual Server | Status |
|------------|---------------|--------|
| `list_teams` | `list-joined-teams` | ❌ MISMATCH |
| `list_channels` | `list-team-channels` | ❌ MISMATCH |
| `send_channel_message` | `send-channel-message` | ❌ MISMATCH |
| `list_chats` | `list-chats` | ❌ MISMATCH |
| `send_chat_message` | `send-chat-message` | ❌ MISMATCH |
| `list_sites` | `search-sharepoint-sites` | ❌ MISMATCH |
| `list_site_files` | `list-sharepoint-site-items` | ❌ MISMATCH |

### 2. CLI Tool Name Mismatches (CRITICAL)

The `ms365_cli.py` script uses incorrect tool names:

- `search-people` → Should be `search-query`
- `verify-login` → Need to verify this exists
- `list-accounts` → Need to verify this exists

### 3. mcporter.example.json Syntax Error (HIGH)

The example JSON uses shell variable substitution syntax `${VAR}` which is invalid JSON:

```json
"MS365_MCP_CLIENT_ID": "${MS365_MCP_CLIENT_ID}"
```

This won't work - JSON doesn't support variable interpolation.

### 4. scripts/check-auth.sh Uses Wrong Tool Name (HIGH)

Line 26: `mcporter call ms365.list_messages` should be `mcporter call ms365.list-mail-messages`

### 5. Missing Tools in Documentation (MEDIUM)

Many available tools are not documented:

**Missing Email Tools:**
- `list-mail-folders`
- `list-mail-folder-messages`
- `create-draft-email`

**Missing Calendar Tools:**
- `get-calendar-view`

**Missing OneDrive Tools:**
- `list-drives`
- `get-drive-root-item`
- `upload-new-file`

**Missing Excel Tools (entire category missing):**
- `list-excel-worksheets`
- `get-excel-range`
- `create-excel-chart`
- `format-excel-range`
- `sort-excel-range`

**Missing OneNote Tools:**
- `create-onenote-page`

**Missing Planner Tools (entire category missing):**
- `list-planner-tasks`
- `get-planner-plan`
- `list-plan-tasks`
- `get-planner-task`
- `create-planner-task`

**Missing Teams/Org Tools:**
- `get-chat`
- `list-chat-messages`
- `get-chat-message`
- `list-chat-message-replies`
- `reply-to-chat-message`
- `get-team`
- `get-team-channel`
- `list-channel-messages`
- `get-channel-message`
- `list-team-members`

**Missing SharePoint Tools:**
- `get-sharepoint-site`
- `get-sharepoint-site-by-path`
- `list-sharepoint-site-drives`
- `get-sharepoint-site-drive-by-id`
- `get-sharepoint-site-item`
- `list-sharepoint-site-lists`
- `get-sharepoint-site-list`
- `list-sharepoint-site-list-items`
- `get-sharepoint-site-list-item`
- `get-sharepoint-sites-delta`

**Missing Shared Mailbox Tools:**
- `list-shared-mailbox-messages`
- `list-shared-mailbox-folder-messages`
- `get-shared-mailbox-message`
- `send-shared-mailbox-mail`

**Missing User Management Tools:**
- `list-users`

### 6. SKILL.md Issues (MEDIUM)

- Uses CLI commands with potentially incorrect tool names
- `search-people` tool doesn't exist (should be `search-query`)
- Missing many available tools
- No mention of Excel or Planner capabilities

### 7. Missing Documentation Features (LOW)

- No mention of `--preset` option for reducing tool load
- No mention of `--discovery` mode
- No mention of `--cloud china` option for 21Vianet
- No mention of Azure Key Vault integration
- No mention of `--enable-auth-tools` for HTTP mode

## Files Requiring Fixes

1. **README.md** - Fix all tool names in reference tables
2. **SKILL.md** - Fix CLI examples and tool references
3. **SKILL.clawdbot.md** - Same fixes as SKILL.md
4. **ms365_cli.py** - Fix tool name mappings
5. **mcporter.example.json** - Fix variable interpolation
6. **scripts/check-auth.sh** - Fix tool name

## Recommendations for Polish Stage

1. **Add a complete tool reference appendix** with all 90+ tools organized by category
2. **Add usage examples** for common workflows
3. **Add troubleshooting section** for specific error codes
4. **Document the difference between personal and org mode** more clearly
5. **Add information about token caching** and where credentials are stored
6. **Document the TOON format** with before/after examples
7. **Add section on tool presets** for performance optimization
