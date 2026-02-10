# Contributing to MS 365 Skill

## Getting Started

1. Fork the repository
2. Clone your fork
3. Create a new branch for your feature or bugfix
4. Make your changes
5. Run tests and ensure everything works
6. Submit a pull request

## Development Setup

1. Ensure you have Python 3.6+ installed (uses only standard library)
2. Install MCP server for testing:
   ```bash
   npm install -g @softeria/ms-365-mcp-server
   ```
3. Set up Azure AD credentials (see README.md for instructions)
4. Test the CLI:
   ```bash
   python3 ms365_cli.py status
   ```

## Contribution Guidelines

- Follow the existing code style
- Add tests for new features
- Update documentation
- Keep pull requests focused and concise

## Reporting Issues

- Use GitHub Issues
- Provide clear, reproducible steps
- Include error messages and context

## Code of Conduct

Be respectful, inclusive, and constructive in all interactions.