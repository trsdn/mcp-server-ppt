# PowerPoint MCP Server Skill

Agent Skill for AI assistants using the PowerPoint MCP Server via the Model Context Protocol.

## Best For

- **Conversational AI** (Claude Desktop, VS Code Chat)
- Exploratory automation with iterative reasoning
- Self-healing workflows needing rich introspection
- Long-running autonomous tasks with continuous context

## Installation

### GitHub Copilot

The [PowerPoint MCP Server VS Code extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.ppt-mcp) installs this skill automatically to `~/.copilot/skills/ppt-mcp/`.

Enable skills in VS Code settings:
```json
{
  "chat.useAgentSkills": true
}
```

### Other Platforms

Extract to your AI assistant's skills directory:

| Platform | Location |
|----------|----------|
| **Claude Code** | `.claude/skills/ppt-mcp/` |
| **Cursor** | `.cursor/skills/ppt-mcp/` |
| **Windsurf** | `.windsurf/skills/ppt-mcp/` |
| **Gemini CLI** | `.gemini/skills/ppt-mcp/` |
| **Codex** | `.codex/skills/ppt-mcp/` |
| **And 36+ more** | Via `npx skills` |
| **Goose** | `.goose/skills/ppt-mcp/` |

Or use npx:
```bash
# Interactive - prompts to select ppt-cli, ppt-mcp, or both
npx skills add sbroenne/mcp-server-ppt

# Or specify directly
npx skills add sbroenne/mcp-server-ppt --skill ppt-mcp
```

## Contents

```
ppt-mcp/
├── SKILL.md           # Main skill definition with MCP tool guidance
├── VERSION            # Version tracking
├── README.md          # This file
└── references/        # Detailed domain-specific guidance
    ├── behavioral-rules.md
    ├── anti-patterns.md
    ├── workflows.md
    ├── range.md
    ├── table.md
    ├── worksheet.md
    ├── chart.md
    ├── slicer.md
    ├── powerquery.md
    ├── datamodel.md
    ├── conditionalformat.md
    └── claude-desktop.md
```

## MCP Server Setup

The skill works with the PowerPoint MCP Server. See [Installation Guide](https://PptMcpserver.dev/installation/) for setup instructions.

## Related

- [PowerPoint CLI Skill](https://github.com/sbroenne/mcp-server-ppt/releases) - For coding agents preferring CLI tools
- [Documentation](https://PptMcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-ppt)
