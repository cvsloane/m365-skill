# Contributing to MS 365 Skill

## Getting Started

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/m365-skill.git
   ```
3. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   ```
4. Make your changes
5. Run tests and ensure everything works
6. Commit your changes with a descriptive commit message
7. Push to your fork and submit a pull request

## Development Setup

### Prerequisites
- Python 3.8+
- npm
- Git

### Installation
1. Install the MCP server:
   ```bash
   npm install -g @softeria/ms-365-mcp-server
   ```
2. Set up an Azure AD app or use device code flow for testing
3. Recommended: Create a virtual environment for Python dependencies

## Contribution Guidelines

### Code Style
- Follow PEP 8 guidelines for Python
- Use consistent indentation (4 spaces)
- Keep lines under 120 characters
- Add type hints where possible

### Testing
- Add unit tests for new features
- Ensure 100% code coverage for added functionality
- Use `python3 -m pytest` for running tests

### Documentation
- Update README.md and SKILL.md for new features
- Add docstrings to functions
- Include example usage in documentation

### Pull Request Process
- Keep PRs focused and concise
- Describe the purpose of your changes in the PR description
- Ensure all CI checks pass
- Be open to code review feedback

## Reporting Issues

- Use GitHub Issues for bug reports and feature requests
- Provide clear, reproducible steps for bugs
- Include:
  - Python and npm versions
  - MCP server version
  - Full error traceback
  - Sample code demonstrating the issue

## Code of Conduct

- Be respectful and inclusive
- Provide constructive feedback
- Collaborate openly and positively

## Contact

For questions about contributing, contact the project maintainers through GitHub Issues.