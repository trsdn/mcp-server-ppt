# PowerPoint CLI Skill

Agent Skill for AI coding assistants using the PowerPoint CLI tool (`pptcli`).

## Best For

- **Coding agents** (GitHub Copilot, Cursor, Windsurf, Codex, Gemini CLI, and 38+ more)
- Token-efficient workflows (no large tool schemas)
- Discoverable via `pptcli --help`
- Scriptable in PowerShell pipelines, CI/CD, batch processing
- Quiet mode (`-q`) outputs clean JSON only

## Why CLI Over MCP?

Modern coding agents increasingly favor CLI-based workflows:

```powershell
# Token-efficient: No schema overhead
pptcli -q session open C:\Data\Report.pptx
pptcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values-json '[["Hello"]]'
pptcli -q session close --session 1 --save
```

## Installation

### GitHub Copilot

The [PowerPoint MCP Server VS Code extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.ppt-mcp) installs this skill automatically to `~/.copilot/skills/ppt-cli/`.

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
| **Claude Code** | `.claude/skills/ppt-cli/` |
| **Cursor** | `.cursor/skills/ppt-cli/` |
| **Windsurf** | `.windsurf/skills/ppt-cli/` |
| **Gemini CLI** | `.gemini/skills/ppt-cli/` |
| **Codex** | `.codex/skills/ppt-cli/` |
| **And 36+ more** | Via `npx skills` |
| **Goose** | `.goose/skills/ppt-cli/` |

Or use npx:
```bash
# Interactive - prompts to select ppt-cli, ppt-mcp, or both
npx skills add sbroenne/mcp-server-ppt

# Or specify directly
npx skills add sbroenne/mcp-server-ppt --skill ppt-cli
```

## Contents

```
ppt-cli/
├── SKILL.md           # Main skill definition with CLI command guidance
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
    └── conditionalformat.md
```

## CLI Installation

Install the CLI tool via NuGet:
```powershell
dotnet tool install --global PptMcp.CLI
```

Verify installation:
```powershell
pptcli --version
pptcli --help
```

## Related

- [PowerPoint MCP Skill](https://github.com/sbroenne/mcp-server-ppt/releases) - For conversational AI (Claude Desktop, VS Code Chat)
- [Documentation](https://PptMcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-ppt)
