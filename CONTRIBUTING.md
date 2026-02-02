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
- **Node.js**: v20+ (for MCP server)
- **Python**: 3.8+ (for CLI wrapper)
- **npm**: v8+ (comes with Node.js)
- **Git**: For version control

### Environment Setup
```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/m365-skill.git
cd m365-skill

# Install MCP server globally
npm install -g @softeria/ms-365-mcp-server

# Verify installation
npx -y @softeria/ms-365-mcp-server --version

# Test basic functionality
python3 ms365_cli.py status
```

### Development Environment
```bash
# Create virtual environment (optional)
python3 -m venv venv
source venv/bin/activate

# Install development dependencies
npm install --save-dev @types/node typescript

# Set up environment variables
cp .env.example .env
# Edit .env with your credentials
```

## Contribution Guidelines

### Code Style
- Follow existing Python and JavaScript conventions
- Use 4 spaces for indentation (no tabs)
- Limit lines to 80 characters where possible
- Use descriptive variable and function names
- Add docstrings to all public functions

### Testing
- Test new features with both device code and Azure AD authentication
- Include error handling tests
- Test edge cases and boundary conditions
- Verify integration with Clawdbot (if applicable)

### Documentation
- Update README.md for new features
- Add API documentation to API_REFERENCE.md
- Update architecture documentation if adding new components
- Include examples in documentation

### Pull Request Guidelines
- Keep pull requests focused and concise
- Include a clear description of changes
- Reference related issues with `#issue-number`
- Ensure all tests pass
- Update documentation as needed

## Development Workflow

### Feature Development
1. Create a feature branch from `main`
2. Make your changes
3. Test thoroughly
4. Update documentation
5. Submit pull request

### Bug Fixes
1. Create a bugfix branch from `main`
2. Fix the issue
3. Add tests to prevent regression
4. Submit pull request

### Documentation Updates
1. Create a docs branch from `main`
2. Update relevant documentation files
3. Verify accuracy of changes
4. Submit pull request

## Testing

### Manual Testing
```bash
# Test authentication
python3 ms365_cli.py login
python3 ms365_cli.py status

# Test email functionality
python3 ms365_cli.py mail list --top 5
python3 ms365_cli.py mail send --to "test@example.com" --subject "Test" --body "Hello"

# Test calendar functionality
python3 ms365_cli.py calendar list --top 5

# Test file functionality
python3 ms365_cli.py files list

# Test task functionality
python3 ms365_cli.py tasks lists

# Test contact functionality
python3 ms365_cli.py contacts list --top 5
```

### Integration Testing
- Test with different authentication methods
- Test with organization mode enabled
- Test with read-only mode
- Test with different output formats

### Error Handling Testing
- Test with invalid credentials
- Test with network failures
- Test with API rate limits
- Test with invalid parameters

## Reporting Issues

### Bug Reports
Use GitHub Issues with the following template:
```markdown
## Bug Description
Brief description of the issue

## Steps to Reproduce
1. Step one
2. Step two
3. Step three

## Expected Behavior
What should happen

## Actual Behavior
What actually happens

## Environment
- OS: [e.g., Ubuntu 20.04]
- Node.js: [e.g., v20.10.0]
- Python: [e.g., 3.8.10]
- m365-skill: [e.g., v1.0.0]

## Error Messages
```

### Feature Requests
Use GitHub Issues with the following template:
```markdown
## Feature Description
Brief description of the requested feature

## Use Case
Why this feature is needed

## Proposed Implementation
How the feature could be implemented

## Alternatives Considered
Other approaches that were considered

## Additional Context
Any other relevant information
```

## Code of Conduct

Be respectful, inclusive, and constructive in all interactions.

### Our Pledge
We commit to making participation in our project a harassment-free experience for everyone, regardless of age, body size, disability, ethnicity, sex characteristics, gender identity and expression, level of experience, education, socio-economic status, nationality, personal appearance, race, religion, or sexual identity and orientation.

### Our Standards
Examples of behavior that contributes to a positive environment for our community:
- Using welcoming and inclusive language
- Being respectful of differing viewpoints and experiences
- Gracefully accepting constructive criticism
- Focusing on what is best for the community
- Showing empathy towards other community members

### Enforcement
Instances of abusive, harassing, or otherwise unacceptable behavior may be reported to the maintainers of the project. All complaints will be reviewed and investigated promptly and fairly.

## License

By contributing to this project, you agree that your contributions will be licensed under the MIT License.

## Support

- **Documentation**: See [README.md](./README.md), [API_REFERENCE.md](./API_REFERENCE.md), and [ARCHITECTURE.md](./ARCHITECTURE.md)
- **Issues**: [GitHub Issues](https://github.com/cvsloane/m365-skill/issues)
- **Discussions**: [GitHub Discussions](https://github.com/cvsloane/m365-skill/discussions)