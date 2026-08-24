# ppt-mcp-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft PowerPoint via the [PowerPoint MCP Server](https://github.com/trsdn/mcp-server-ppt).

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate PowerPoint through 223 MCP operations:

- **Presentation management** — open, create, save, close
- **Slides & layouts** — create, duplicate, move, apply layouts
- **Shapes & text** — geometry, fill, effects, formatting, bullets
- **Tables & charts** — create, populate, format
- **Design & themes** — palettes, archetypes, layout grids
- **Export** — PDF, images, video
- **VBA macros, accessibility audits**, and more

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
