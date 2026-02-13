# Contributing to MS 365 Skill

Thank you for your interest in improving this project! This guide will help you get started.

---

## Quick Start

```bash
# 1. Fork the repository on GitHub
# 2. Clone your fork
git clone https://github.com/YOUR_USERNAME/m365-skill.git
cd m365-skill

# 3. Create a feature branch
git checkout -b feature/your-feature-name

# 4. Make your changes
# 5. Test thoroughly
# 6. Submit a pull request
```

---

## Development Setup

### Requirements

- **Python 3.8+** — for the CLI wrapper
- **Node.js & npm** — for the MCP server

### Install Dependencies

```bash
# Install the MCP server globally
npm install -g @softeria/ms-365-mcp-server

# No Python dependencies needed!
# The CLI uses only standard library modules
```

### Configure for Local Development

To test your changes, you'll need Microsoft 365 credentials:

1. **Option A:** Set up an Azure AD app (see [README.md](./README.md#option-a-azure-ad-app-registration-recommended))
2. **Option B:** Use device code flow for personal testing

Set environment variables:
```bash
export MS365_MCP_CLIENT_ID="your-client-id"
export MS365_MCP_CLIENT_SECRET="your-secret"
export MS365_MCP_TENANT_ID="consumers"
```

---

## How to Contribute

### Reporting Bugs

When opening an issue, please include:

- **Clear description** of what went wrong
- **Steps to reproduce** the problem
- **Expected vs. actual behavior**
- **Error messages** (full stack traces if available)
- **Environment details:**
  - Operating system
  - Python version (`python3 --version`)
  - npm package version (`npm list -g @softeria/ms-365-mcp-server`)

### Suggesting Features

We welcome feature suggestions! Please:

- Check existing issues first
- Describe the use case clearly
- Explain why the feature would benefit users
- Consider implementation complexity

### Code Contributions

#### Pull Request Process

1. **Keep PRs focused** — one feature or fix per pull request
2. **Update documentation** — README, SKILL.md, or CLAUDE.md as needed
3. **Test your changes** — verify commands work as expected
4. **Follow existing style** — match the code patterns you see

#### Code Style

- **Python:** Follow PEP 8 conventions
- **Documentation:** Use clear, concise language
- **Commit messages:** Write descriptive commit messages
  - `feat:` for new features
  - `fix:` for bug fixes
  - `docs:` for documentation changes
  - `refactor:` for code restructuring
  - `chore:` for maintenance tasks

#### Documentation Guidelines

When updating documentation:

- **README.md** — Focus on user setup and troubleshooting
- **SKILL.md** — Focus on Clawdbot behavior and command reference
- **CLAUDE.md** — Focus on developer setup and architecture
- Use tables for structured data (options, environment variables)
- Include examples for complex commands

---

## Project Structure

```
m365-skill/
├── README.md              # User-facing documentation
├── SKILL.md               # Clawdbot skill definition
├── CLAUDE.md              # Developer guidance (this file)
├── CONTRIBUTING.md        # Contribution guidelines
├── LICENSE.md             # MIT License
├── ms365_cli.py           # Python CLI wrapper
├── mcporter.example.json  # Example mcporter config
├── project_status.md      # Auto-generated status file
└── scripts/               # Helper scripts
    ├── auth-device.sh     # Device code authentication
    ├── check-auth.sh      # Check auth status
    ├── list-tools.sh      # List available tools
    └── start-server.sh    # Start MCP server
```

---

## Testing Checklist

Before submitting a PR, verify:

- [ ] `python3 ms365_cli.py login` works (if testing auth)
- [ ] `python3 ms365_cli.py status` returns expected output
- [ ] `python3 ms365_cli.py mail list` retrieves emails
- [ ] `python3 ms365_cli.py calendar list` retrieves events
- [ ] Documentation reflects your changes
- [ ] No syntax errors in Python or Markdown files

---

## Code of Conduct

This project strives to be welcoming and inclusive:

- **Be respectful** in all interactions
- **Be constructive** when providing feedback
- **Be inclusive** — welcome contributors of all backgrounds
- **Assume good intent** — we're all here to improve the project

Harassment, discrimination, or toxic behavior will not be tolerated.

---

## Questions?

- Open an issue for bugs or feature requests
- Check existing documentation first
- Be patient — maintainers are volunteers

Thank you for contributing!
