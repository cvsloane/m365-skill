# Contributing to MS 365 Skill

## Getting Started

1. Fork the repository
2. Clone your fork
3. Create a new branch for your feature or bugfix
4. Make your changes
5. Run tests and ensure everything works
6. Submit a pull request

## Development Setup

### Prerequisites

- Python 3.8+ (for CLI wrapper)
- Node.js 16+ (for MCP server)
- Git
- Microsoft 365 account (for testing)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/cvsloane/m365-skill.git
   cd m365-skill
   ```

2. Install the MCP server globally:
   ```bash
   npm install -g @softeria/ms-365-mcp-server
   ```

3. Install Python dependencies (if modifying CLI):
   ```bash
   pip install -r requirements.txt  # If requirements.txt exists
   ```

### Testing Setup

1. Set up Azure AD app registration (see README.md for detailed instructions)
2. Configure environment variables:
   ```bash
   export MS365_MCP_CLIENT_ID=your-client-id
   export MS365_MCP_CLIENT_SECRET=your-client-secret
   export MS365_MCP_TENANT_ID=your-tenant-id
   ```

3. Test the installation:
   ```bash
   # Test MCP server
   mcporter list ms365
   
   # Test CLI wrapper
   python3 ms365_cli.py status
   
   # Test authentication
   python3 ms365_cli.py login
   ```

## Development Workflow

### Code Style

- Follow PEP 8 for Python code
- Use descriptive variable and function names
- Add docstrings to all public functions
- Keep functions focused and single-purpose

### Testing

- Test all new functionality with real Microsoft 365 accounts
- Include both positive and negative test cases
- Test with both personal and organizational accounts (when possible)
- Verify error handling and edge cases

### Documentation

- Update README.md for any new features
- Update SKILL.md for new CLI commands
- Add comments to complex code sections
- Update this CONTRIBUTING.md if adding new development processes

## Project Structure

```
m365-skill/
├── README.md              # User-facing documentation
├── SKILL.md               # Clawbot skill definition
├── SKILL.clawdbot.md      # Clawbot-specific documentation
├── CONTRIBUTING.md        # This file
├── CLAUDE.md              # Claude Code guidance
├── ms365_cli.py           # Python CLI wrapper
├── mcporter.example.json  # Example mcporter configuration
├── scripts/               # Helper scripts
│   ├── auth-device.sh     # Device code authentication
│   ├── check-auth.sh      # Authentication status check
│   ├── list-tools.sh      # List available tools
│   └── start-server.sh    # Start HTTP server
└── project_status.md      # Development status tracking
```

## Available Scripts

- `scripts/auth-device.sh` - Interactive device code authentication
- `scripts/check-auth.sh` - Check authentication status
- `scripts/list-tools.sh` - List available MCP tools
- `scripts/start-server.sh` - Start MCP server in HTTP mode

## Contribution Guidelines

- Follow the existing code style
- Add tests for new features
- Update documentation
- Keep pull requests focused and concise
- Ensure all existing functionality continues to work
- Test with both personal and organizational accounts when applicable

## Reporting Issues

- Use GitHub Issues
- Provide clear, reproducible steps
- Include error messages and context
- Specify whether you're using personal or organizational account
- Include your environment (OS, Node.js version, Python version)

## Pull Request Process

1. Create a feature branch from `main`
2. Make your changes
3. Add tests if applicable
4. Update documentation
5. Ensure all existing tests pass
6. Submit a pull request with a clear description
7. Link to any related issues

## Code of Conduct

Be respectful, inclusive, and constructive in all interactions.

## Support

- For MCP server issues: https://github.com/Softeria/ms-365-mcp-server/issues
- For skill-specific issues: https://github.com/cvsloane/m365-skill/issues