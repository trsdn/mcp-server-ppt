---
name: MCP Server Issue
about: Report issues with the MCP Server for AI assistants
title: '[MCP] '
labels: 'mcp-server'
assignees: ''

---

## Issue Description
A clear and concise description of the MCP Server issue.

## AI Assistant
Which AI assistant are you using with the MCP Server?
- [ ] **GitHub Copilot** (VS Code, Visual Studio, etc.)
- [ ] **Claude Desktop** (Anthropic)
- [ ] **ChatGPT** (OpenAI)
- [ ] **Other**: [please specify]

## MCP Tool & Action
Which MCP tool and action are experiencing issues?
- **Tool**: [e.g., slide, shape, text, chart, vba, file]
- **Action**: [e.g., list, view, import, export, update, refresh, delete, etc.]
- **File Path**: [e.g., "C:\Data\presentation.pptx"]
- **Additional Parameters**: [describe any other parameters used]

## Expected Behavior
What did you expect the MCP Server to do?

## Actual Behavior
What did the MCP Server actually do?

## Error Response
If you received an error, paste the full JSON response:
```json
{
  "error": "paste error here"
}
```

## MCP Server Configuration
How is the MCP Server configured?

**Configuration file location**: [e.g., `.config/Code/User/globalStorage/github.copilot-chat/config.json`]

**MCP Configuration**:
```json
{
  "mcpServers": {
    "ppt-mcp": {
      "command": "mcp-ppt",
      // or other configuration
    }
  }
}
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **PowerPoint Version**: [e.g. PowerPoint 365, PowerPoint 2019]
- **PptMcp Version**: [e.g. v1.0.0 - run `mcp-ppt --version` or `dotnet tool list -g`]
- **.NET Version**: [Run `dotnet --version`]
- **Installation Method**: 
  - [ ] Global .NET tool (`dotnet tool install --global PptMcp.McpServer`)
  - [ ] Source build
  - [ ] Other: [please specify]

## MCP Server Logs
If possible, provide relevant logs from the MCP Server:
```
[Paste logs here]
```

## Steps to Reproduce
1. Configure AI assistant with MCP Server
2. Ask AI assistant: "..."
3. MCP Server receives request for tool: [tool_name], action: [action_name]
4. See error

## Conversation Context (Optional)
If helpful, provide the conversation you had with the AI assistant that led to this issue:
```
User: "Can you list all slides in my presentation?"
AI: [response]
[MCP Server error occurs]
```

## PowerPoint File Details
- **File Format**: [.pptx or .pptm]
- **File Size**: [approximate size]
- **Contains**: 
  - [ ] Power Queries
  - [ ] VBA Macros
  - [ ] Named Ranges
  - [ ] Multiple slides
  - [ ] External connections

## VBA-Related Issues (if applicable)
- [ ] VBA trust is properly configured (`PptMcp check-vba-trust`)
- [ ] Using .pptm file format for VBA operations
- [ ] VBA module exists in the presentation
- [ ] Macro security settings allow programmatic access

## Additional Context
Add any other context about the problem here, including:
- Screenshots of AI assistant interaction
- Sample PowerPoint files (with sensitive data removed)
- Other relevant information
