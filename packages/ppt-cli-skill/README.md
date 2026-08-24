# ppt-cli-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft PowerPoint via the [pptcli](https://github.com/trsdn/mcp-server-ppt) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate PowerPoint from scripts and CI/CD pipelines:

- **Presentation management** — open, create, save, close
- **Slides & layouts** — create, duplicate, move, apply layouts
- **Shapes & text** — geometry, fill, effects, formatting, bullets
- **Tables & charts** — create, populate, format
- **Design & themes** — palettes, archetypes, layout grids
- **Export** — PDF, images, video
- **VBA macros, accessibility audits**, and more

## Requirements

- Windows with Microsoft PowerPoint 2016+ installed
- Install the CLI: `dotnet tool install --global PptMcp.CLI`

## Install

```bash
npx skillpm install ppt-cli-skill
```

Or with npm directly:

```bash
npm install ppt-cli-skill
```

## License

MIT
