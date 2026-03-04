# ppt-mcp-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft PowerPoint via the [PowerPoint MCP Server](https://PptMcpserver.dev).

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate PowerPoint through 225 MCP operations:

- **Presentation management** — open, create, save, close
- **Range operations** — read/write values, formatting, formulas
- **Tables & PivotTables** — create, modify, refresh
- **Charts** — create and configure chart types
- **Power Query (M code)** — create and edit queries
- **Data Model (DAX)** — add measures and calculated columns
- **Conditional formatting, slicers, VBA macros**, and more

## Requirements

- Windows with Microsoft PowerPoint 2016+ installed
- [PowerPoint MCP Server](https://github.com/trsdn/mcp-server-ppt) running

## Install

```bash
npx skillpm install ppt-mcp-skill
```

Or with npm directly:

```bash
npm install ppt-mcp-skill
```

## License

MIT
