# MS 365 Skill Architecture

## Overview

The MS 365 Skill is a Clawdbot skill that provides Microsoft 365 integration through the Microsoft Graph API. It uses a layered architecture with clear separation of concerns.

## Architecture Components

### 1. Skill Layer (Clawdbot Integration)

**File**: `SKILL.md`, `SKILL.clawdbot.md`

**Responsibilities:**
- Define skill activation triggers
- Provide natural language processing for user requests
- Map user intent to appropriate CLI commands
- Handle Clawdbot-specific configuration

**Key Components:**
- Activation keywords for different Microsoft 365 services
- Command mapping logic
- User interaction prompts
- Authentication state management

### 2. CLI Interface Layer

**File**: `ms365_cli.py`

**Responsibilities:**
- Provide command-line interface for direct interaction
- Handle argument parsing and validation
- Manage MCP server communication
- Format and display output

**Key Components:**
- `call_mcp()`: Core MCP server communication
- Command handlers for each service (mail, calendar, files, etc.)
- Output formatting utilities
- Authentication management

### 3. MCP Server Layer

**Package**: `@softeria/ms-365-mcp-server`

**Responsibilities:**
- Implement Microsoft Graph API client
- Handle OAuth authentication flows
- Provide standardized tool interface
- Manage API rate limiting and error handling

**Key Features:**
- Device code authentication
- Azure AD app authentication
- Tool discovery and execution
- Response formatting

### 4. Microsoft Graph API Layer

**Service**: Microsoft Graph API

**Responsibilities:**
- Provide access to Microsoft 365 services
- Handle API authentication and authorization
- Manage data operations (CRUD)
- Enforce security and compliance policies

## Data Flow

### 1. User Request Flow
```
User Request → Clawbot Skill → CLI Interface → MCP Server → Graph API → Microsoft 365 Service
```

### 2. Authentication Flow
```
User Login → Device Code/Azure AD → Token Cache → MCP Server → Graph API
```

### 3. Data Processing Flow
```
Graph API Response → MCP Server → CLI Interface → Formatted Output → User Display
```

## Service Integration

### Email (Outlook)
- **API Endpoint**: `/me/messages`
- **Operations**: List, read, send, reply, forward, delete, search
- **Authentication**: Delegated permissions (Mail.ReadWrite, Mail.Send)

### Calendar
- **API Endpoint**: `/me/events`
- **Operations**: List, create, update, delete, search
- **Authentication**: Delegated permissions (Calendars.ReadWrite)

### OneDrive
- **API Endpoint**: `/me/drive`
- **Operations**: List, upload, download, create, delete, search
- **Authentication**: Delegated permissions (Files.ReadWrite)

### To Do Tasks
- **API Endpoint**: `/me/todo`
- **Operations**: List, create, update, complete, delete
- **Authentication**: Delegated permissions (Tasks.ReadWrite)

### Contacts
- **API Endpoint**: `/me/contacts`
- **Operations**: List, search, create, read
- **Authentication**: Delegated permissions (Contacts.Read)

### OneNote
- **API Endpoint**: `/me/notes`
- **Operations**: List notebooks, sections, pages, read content
- **Authentication**: Delegated permissions (Notes.Read)

### Organization Features (Teams, SharePoint)
- **API Endpoints**: `/me/chats`, `/me/teams`, `/sites`
- **Operations**: List, send messages, access documents
- **Authentication**: Application permissions (Chat.ReadWrite, Sites.Read.All)
- **Mode**: Requires `--org-mode` flag and Azure AD app registration

## Configuration Architecture

### Environment Variables
```
MS365_MCP_CLIENT_ID        # Azure AD app client ID
MS365_MCP_CLIENT_SECRET    # Azure AD app client secret
MS365_MCP_TENANT_ID        # Azure AD tenant ID
MS365_MCP_ORG_MODE         # Enable organization features
MS365_MCP_READ_ONLY        # Disable write operations
MS365_MCP_OUTPUT_FORMAT    # Response format (toon, etc.)
MS365_MCP_PORT             # HTTP server port
```

### McPorter Configuration
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

## Security Architecture

### Authentication Methods
1. **Device Code Flow**: Interactive, suitable for development
2. **Azure AD App**: Headless, suitable for production/Docker

### Token Management
- Tokens are cached locally after successful authentication
- Refresh tokens are used for automatic renewal
- Token expiration is handled gracefully

### Permission Model
- **Delegated Permissions**: User acts on behalf of themselves
- **Application Permissions**: App acts on behalf of users (org mode)
- **Principle of Least Privilege**: Only required permissions are requested

## Error Handling Architecture

### Error Categories
1. **Authentication Errors**: Invalid credentials, expired tokens
2. **Network Errors**: Connection timeouts, API rate limits
3. **API Errors**: Invalid parameters, permission denied
4. **Configuration Errors**: Missing environment variables

### Error Recovery
- Automatic token refresh when possible
- Graceful degradation for non-critical operations
- Clear error messages with suggested solutions

## Performance Considerations

### Caching Strategy
- Authentication tokens are cached
- API responses are not cached (real-time data)
- File metadata may be cached for performance

### Rate Limiting
- Microsoft Graph API rate limits are respected
- Batch operations are used where possible
- Error handling for rate limit exceeded

## Deployment Architecture

### Development Environment
- Device code authentication
- Local development server
- Direct CLI usage for testing

### Production Environment
- Azure AD app authentication
- Docker container support
- HTTP server mode for remote access
- Environment variable configuration

### Scalability
- Stateless design (except for token cache)
- Horizontal scaling through multiple instances
- Load balancing supported via HTTP mode

## Monitoring and Logging

### Logging Levels
- **Info**: Successful operations, authentication events
- **Warning**: Retryable errors, deprecation notices
- **Error**: Critical failures, authentication issues

### Metrics
- API call success/failure rates
- Authentication success/failure rates
- Response time metrics
- Error type distribution

## Future Architecture Considerations

### Potential Enhancements
1. **Caching Layer**: Add response caching for frequently accessed data
2. **Batch Operations**: Optimize multiple API calls into single requests
3. **Webhooks**: Real-time notifications for calendar events, emails
4. **Multi-tenant Support**: Support for multiple Azure AD tenants
5. **Advanced Filtering**: More sophisticated search and filtering capabilities

### Migration Path
- Current architecture supports incremental improvements
- Backward compatibility maintained for existing APIs
- Configuration-driven feature enabling