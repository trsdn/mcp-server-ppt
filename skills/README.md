# PowerPoint MCP Server - Agent Skills

Two skill packages for AI coding assistants:

| Skill | Target | Best For |
|-------|--------|----------|
| **[ppt-cli](ppt-cli/SKILL.md)** | CLI Tool | Coding agents - token-efficient, `--help` discoverable |
| **[ppt-mcp](ppt-mcp/SKILL.md)** | MCP Server | Conversational AI - rich tool schemas |

## Installation

```bash
# Via VS Code extension (auto-installs ppt-mcp)
# Or via npx:
npx skills add trsdn/mcp-server-ppt --skill ppt-cli   # Coding agents
npx skills add trsdn/mcp-server-ppt --skill ppt-mcp   # Conversational AI
```

## Building

```powershell
dotnet build -c Release
```

Generates `SKILL.md` and copies `shared/` references into each skill's `references/` folder.

## Structure

```
skills/
├── shared/          # Shared behavioral guidance (source of truth)
├── ppt-mcp/       # MCP Server skill (SKILL.md + references/)
├── ppt-cli/       # CLI skill (SKILL.md + references/)
├── templates/       # Scriban templates for SKILL.md generation
├── CLAUDE.md        # Claude Code project instructions
└── .cursorrules     # Cursor-specific rules
```
