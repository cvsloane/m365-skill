# MS 365 Skill Architecture

## Overview

The MS 365 Skill is a Clawdbot integration that provides access to Microsoft 365 services through the Microsoft Graph API. The architecture follows a layered approach with clear separation of concerns.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                    User Interface                           │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │  Clawdbot   │  │  CLI Tool   │  │  Web UI     │        │
│  │  Agent      │  │  (Python)   │  │  (Future)   │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   Clawdbot Integration                      │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │   mcporter  │  │   SKILL.md  │  │   Prompts   │        │
│  │   Bridge    │  │  Definition │  │   Logic     │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   Communication Layer                       │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │  JSON-RPC   │  │   HTTP      │  │   Stdio     │        │
│  │  Protocol   │  │   Server    │  │   Transport │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   MCP Server Layer                          │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │ @softeria/  │  │   Tool      │  │   Auth      │        │
│  │ ms-365-mcp- │  │   Registry  │  │   Manager   │        │
│  │ server      │  │             │  │             │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   Microsoft Graph API                       │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │   REST API  │  │   OAuth 2.0 │  │   Token    │        │
│  │   Endpoint  │  │   Auth      │  │   Cache     │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   Microsoft 365 Services                    │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │   Outlook   │  │   OneDrive  │  │   Calendar  │        │
│  │   (Email)   │  │   (Files)   │  │   (Events)  │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │   To Do     │  │  Contacts   │  │   OneNote   │        │
│  │   (Tasks)   │  │             │  │   (Notes)   │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐        │
│  │    Teams    │  │  SharePoint │  │   Other     │        │
│  │  (Chat)     │  │  (Sites)    │  │  Services   │        │
│  └─────────────┘  └─────────────┘  └─────────────┘        │
└─────────────────────────────────────────────────────────────┘
```

## Component Details

### 1. User Interface Layer

**Clawdbot Agent**
- Natural language processing for user requests
- Intent recognition for MS 365 actions
- Integration with Clawdbot's skill system

**CLI Tool (`ms365_cli.py`)**
- Command-line interface for direct interaction
- Python wrapper around MCP server
- Standalone testing and debugging

### 2. Clawbot Integration Layer

**mcporter Bridge**
- JSON-RPC communication with MCP server
- Environment variable management
- Process spawning and I/O handling

**SKILL.md Definition**
- Clawbot skill metadata
- Activation triggers and configuration
- Usage examples and prompts

### 3. Communication Layer

**JSON-RPC Protocol**
- Standardized request/response format
- Tool call and result handling
- Error reporting and diagnostics

**HTTP Server Mode**
- Alternative to stdio transport
- Better for containerized deployments
- REST-like API access

### 4. MCP Server Layer

**@softeria/ms-365-mcp-server**
- Microsoft Graph API abstraction
- Tool registry and execution
- Authentication and token management

**Tool Registry**
- Dynamic tool discovery
- Parameter validation and transformation
- Result formatting and pagination

### 5. Microsoft Graph API Layer

**REST API Endpoint**
- Direct communication with Microsoft services
- OAuth 2.0 authentication flow
- Rate limiting and error handling

**Token Management**
- OAuth token caching and refresh
- Device code flow support
- Azure AD app authentication

### 6. Microsoft 365 Services Layer

**Core Services**
- **Outlook**: Email, calendar, contacts
- **OneDrive**: File storage and management
- **To Do**: Task management
- **OneNote**: Note taking and organization

**Organization Services** (with --org-mode)
- **Microsoft Teams**: Chat and channel messaging
- **SharePoint**: Site and document management
- **Other**: Additional enterprise services

## Data Flow

### Authentication Flow
1. User initiates authentication (device code or Azure AD)
2. MCP server handles OAuth flow
3. Tokens are cached for future requests
4. Subsequent requests use cached tokens

### Request Flow
1. User makes request via Clawbot or CLI
2. Request is parsed and validated
3. MCP tool is called with appropriate parameters
4. Microsoft Graph API request is made
5. Response is processed and formatted
6. Result is returned to user

### Error Handling Flow
1. Error occurs at any layer
2. Error is captured and formatted
3. Appropriate error message is generated
4. User is informed of the issue
5. Recovery options are suggested

## Configuration Management

### Environment Variables
- Azure AD credentials for headless operation
- Server mode configuration
- Feature flags (org-mode, read-only, etc.)

### Configuration Files
- `mcporter.json`: Server bridge configuration
- `SKILL.md`: Clawbot skill definition
- `scripts/`: Helper scripts for common operations

## Security Considerations

### Authentication
- OAuth 2.0 with PKCE for device flow
- Azure AD app registration for production
- Token encryption and secure storage

### Data Access
- Principle of least privilege
- Read-only mode option
- Audit logging capabilities

### Network Security
- HTTPS for all API communications
- Firewall rules for server mode
- Input validation and sanitization

## Deployment Architecture

### Development Environment
- Local testing with device code flow
- Direct MCP server interaction
- Debug logging enabled

### Production Environment
- Azure AD app authentication
- Containerized deployment
- Load balancing and scaling

### Monitoring
- Health checks and diagnostics
- Performance metrics
- Error tracking and alerting

## Extension Points

### Custom Tools
- Additional Microsoft Graph endpoints
- Custom business logic
- Integration with other services

### Authentication Methods
- Multi-factor authentication
- SSO integration
- Certificate-based auth

### Output Formats
- Custom response formatting
- Localization support
- Accessibility features

## Performance Considerations

### Caching
- Token caching to reduce auth overhead
- Response caching for frequent queries
- Result pagination for large datasets

### Optimization
- Batch operations where possible
- Lazy loading of large responses
- Connection pooling for API calls

### Scalability
- Stateless design for horizontal scaling
- Rate limiting and throttling
- Resource management for concurrent requests