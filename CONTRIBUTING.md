# Contributing to MS 365 Skill

Thank you for your interest in improving the Microsoft 365 Skill for Clawdbot! This document will help you get started contributing effectively.

---

## Getting Started

### Prerequisites

- **Python 3.8+** — for the CLI wrapper
- **Node.js 16+** — for the MCP server
- **Git** — for version control
- **A Microsoft 365 account** — for testing (personal or work)

### Development Setup

1. **Fork the repository** on GitHub

2. **Clone your fork:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/m365-skill.git
   cd m365-skill
   ```

3. **Install the MCP server:**
   ```bash
   npm install -g @softeria/ms-365-mcp-server
   ```

4. **Set up authentication** (choose one):
   - **Quick test:** Use device code flow (no Azure setup needed)
   - **Full development:** [Set up Azure AD app](./README.md#option-a-azure-ad-app-recommended)

5. **Test your setup:**
   ```bash
   python3 ms365_cli.py status
   python3 ms365_cli.py mail list --top 5
   ```

---

## Development Workflow

### 1. Create a Branch

```bash
git checkout -b feature/your-feature-name
# or
git checkout -b fix/issue-description
```

Use descriptive branch names:
- `feature/add-contact-creation`
- `fix/calendar-timezone-handling`
- `docs/improve-troubleshooting`

### 2. Make Your Changes

**When adding features:**
- Update `ms365_cli.py` with new command handlers
- Add corresponding argument parsers
- Update both `SKILL.md` and `SKILL.clawdbot.md`
- Update `README.md` if user-facing

**When fixing bugs:**
- Include a test case that reproduces the issue
- Verify the fix doesn't break existing functionality

**When updating docs:**
- Keep tone consistent with existing documentation
- Use clear, concise language
- Include examples where helpful

### 3. Test Thoroughly

```bash
# Test authentication
python3 ms365_cli.py status

# Test affected commands
python3 ms365_cli.py mail list
python3 ms365_cli.py calendar list
# etc.

# Test error handling
echo "Test invalid inputs"
```

### 4. Commit Your Changes

Use clear, conventional commit messages:

```bash
# Features
git commit -m "feat: add contact creation command"

# Bug fixes
git commit -m "fix: handle missing timezone in calendar events"

# Documentation
git commit -m "docs: add Azure AD troubleshooting section"

# Refactoring
git commit -m "refactor: extract common API call logic"
```

### 5. Push and Submit PR

```bash
git push origin feature/your-feature-name
```

Then open a pull request on GitHub with:
- Clear description of changes
- Reference to any related issues
- Testing notes
- Screenshots (if UI-related)

---

## Contribution Guidelines

### Code Style

**Python:**
- Follow PEP 8
- Use type hints where appropriate
- Keep functions focused and under 50 lines when possible
- Add docstrings for public functions

**Shell Scripts:**
- Use `set -euo pipefail` for safety
- Quote all variables: `"${VAR}"`
- Add comments for non-obvious logic

**Documentation:**
- Use sentence case for headings
- Keep line lengths reasonable (≤100 chars)
- Use code blocks with language tags

### Testing Requirements

Before submitting:
- [ ] All new code paths have been tested
- [ ] Authentication still works
- [ ] Error messages are helpful
- [ ] Documentation is updated
- [ ] No breaking changes (or clearly documented)

### What to Contribute

**High Priority:**
- Bug fixes
- Documentation improvements
- New Microsoft 365 service integrations
- Error handling enhancements

**Welcome Contributions:**
- Additional CLI commands
- Helper scripts
- Usage examples
- Translations

**Please Discuss First:**
- Breaking changes to existing commands
- New authentication methods
- Major architectural changes

---

## Reporting Issues

### Before You Report

- Check [existing issues](https://github.com/cvsloane/m365-skill/issues)
- Verify you're using the latest version
- Test with a fresh authentication flow

### Creating a Good Issue

**For bugs, include:**
1. Clear description of expected vs. actual behavior
2. Steps to reproduce
3. Environment details (OS, Python version, Node version)
4. Error messages (full traceback if applicable)
5. Authentication method used (device code vs. Azure AD)

**Example:**
```markdown
**Bug:** Calendar events show wrong timezone

**Expected:** Events display in user's local timezone
**Actual:** All times show as UTC

**Steps to Reproduce:**
1. Set timezone to America/New_York
2. Create event: `calendar create --start "2026-02-15T10:00:00"`
3. List events: `calendar list`

**Environment:**
- OS: Ubuntu 22.04
- Python: 3.10
- Node: 18.x
- Auth: Azure AD

**Error:** No error message, incorrect timezone in output
```

**For feature requests, include:**
1. Use case description
2. Proposed solution
3. Alternative approaches considered
4. Willingness to contribute the feature

---

## Code of Conduct

### Our Standards

**Be respectful:** Treat everyone with respect, regardless of experience level.

**Be constructive:** Provide helpful feedback and suggestions.

**Be inclusive:** Welcome newcomers and help them learn.

**Be patient:** Remember that maintainers are volunteers.

### Unacceptable Behavior

- Harassment or discrimination
- Trolling or insulting comments
- Publishing private information
- Disruptive behavior

### Enforcement

Violations may result in temporary or permanent exclusion from the project.

---

## Questions?

- **General questions:** Open a [GitHub Discussion](https://github.com/cvsloane/m365-skill/discussions)
- **Security issues:** Email the maintainer directly (see profile)
- **Quick chat:** Join our community (if applicable)

---

Thank you for contributing! 🎉
