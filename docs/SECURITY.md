# Security Policy

## Supported Versions

We currently support the following versions of PptMcp with security updates:

| Version | Supported          |
| ------- | ------------------ |
| 1.x.x   | :white_check_mark: |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability in PptMcp, please report it responsibly.

### How to Report

1. **Do NOT create a public GitHub issue** for security vulnerabilities
2. Contact me directly
3. Include the following information:
   - Description of the vulnerability
   - Steps to reproduce the issue
   - Potential impact
   - Suggested fix (if you have one)

### What to Expect

- We will acknowledge receipt of your vulnerability report within 48 hours
- We will provide an estimated timeline for addressing the vulnerability within 1 week
- We will notify you when the vulnerability has been fixed
- We will credit you in the security advisory (if you wish)

## Security Considerations

### Enhanced Security Features (Latest Version)

PptMcp implements comprehensive security measures:

- **Input Validation**: All file paths validated with length limits (32767 chars) and extension restrictions
- **File Size Limits**: 1GB maximum file size to prevent DoS attacks  
- **Path Security**: `Path.GetFullPath()` prevents path traversal attacks
- **Resource Protection**: Protection against memory exhaustion and process hanging
- **Code Analysis**: Enhanced security rules enforced (CA2100, CA3003, CA3006, etc.)
- **Quality Enforcement**: All warnings treated as errors for robust code

### PowerPoint COM Automation

PptMcp uses PowerPoint COM automation with security safeguards:

- **Macro Execution**: PptMcp can execute VBA macros when using script-run command
- **VBA Trust**: VBA operations require "Trust access to the VBA project object model" to be manually enabled in PowerPoint settings (one-time setup)
- **File Validation**: Strict file extension validation (.pptx, .pptm, .ppt only)
- **File Access**: PptMcp requires read/write access to PowerPoint files with size validation
- **Process Isolation**: Each command runs in a separate process that terminates after completion
- **PowerPoint Instance**: Creates temporary PowerPoint instances that are properly cleaned up
- **Input Sanitization**: All arguments validated for length and content

### Power Query Privacy Levels

PptMcp implements security-first privacy level handling:

- **Explicit Consent**: Privacy levels must be specified explicitly via `--privacy-level` parameter or `PPT_DEFAULT_PRIVACY_LEVEL` environment variable
- **No Auto-Application**: Privacy levels are never applied automatically without user consent
- **Privacy Detection**: Analyzes existing queries to recommend appropriate privacy levels
- **Clear Guidance**: Provides detailed explanations of privacy level implications
- **Security Options**: Supports None, Private (most secure), Organizational, and Public levels

### VBA Security Considerations

- **Macro Content**: VBA scripts imported via script-import will be executed when called
- **Manual Trust Setup**: VBA trust must be enabled manually through PowerPoint's Trust Center settings (never modified automatically by PptMcp)
- **File Format**: Only .pptm files can contain and execute VBA code
- **Code Injection**: Always validate VBA source files before importing
- **User Control**: PptMcp never modifies registry settings or security configurations automatically

### Best Practices for Users

1. **File Validation**: Only run PptMcp on trusted PowerPoint files
2. **VBA Source Control**: Validate VBA code files before importing with script-import
3. **Network Files**: Be cautious when processing files from network locations
4. **Permissions**: Run PptMcp with minimal necessary permissions
5. **Backup**: Always backup important PowerPoint files before processing
6. **VBA Trust**: Only enable VBA trust in PowerPoint settings on systems where it's needed (manual one-time setup)
7. **Code Review**: Review VBA scripts before execution, especially from external sources
8. **Privacy Levels**: Choose appropriate Power Query privacy levels based on data sensitivity (Private for sensitive data, Organizational for internal data, Public for public APIs)
9. **Environment Variables**: Use `PPT_DEFAULT_PRIVACY_LEVEL` environment variable for consistent automation security

### Known Limitations

- **Windows Only**: PptMcp only works on Windows with PowerPoint installed
- **COM Dependencies**: Relies on PowerPoint COM objects which may have their own security considerations
- **File System Access**: Requires appropriate file system permissions for PowerPoint file access

## Dependency Security

PptMcp has minimal dependencies to reduce attack surface:

- **.NET 10**: Microsoft-maintained runtime with regular security updates
- **Spectre.Console**: Well-maintained library for console output
- **No External APIs**: No network connections or external service dependencies

## Version Updates

- Security patches will be released as soon as possible
- Users are encouraged to keep PptMcp updated to the latest version
- Breaking changes will be clearly documented in release notes

## Contact

For security-related questions or concerns, please contact Torsten Mahr through GitHub issues (for non-sensitive matters) or the security reporting method outlined above.
