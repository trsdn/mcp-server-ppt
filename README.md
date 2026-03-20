# PptMcp - MCP Server for Microsoft PowerPoint

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9-blue.svg)](https://dotnet.microsoft.com/download/dotnet/9.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/trsdn/mcp-server-ppt)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate PowerPoint with AI — A Model Context Protocol (MCP) server for comprehensive PowerPoint automation through conversational AI.**

**MCP Server for PowerPoint** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate PowerPoint through natural language commands. Manage slides, shapes, text, charts, tables, animations, transitions, VBA macros, and more (33 tools with 204 operations).

For multi-phase build / verify / repair workflows from source, the repo also includes the official orchestration client under `src\PptMcp.Agent`.

**🛡️ 100% Safe — Uses PowerPoint's Native COM API** — Zero risk of file corruption. Uses PowerPoint's official COM API ensuring complete safety and compatibility.

**💡 Interactive Development** — See results instantly in PowerPoint. Add slides, create charts, format text, and iterate. PowerPoint becomes your AI-powered workspace.

**Technical Requirements:**
- ⚠️ **Windows Only** — COM interop is Windows-specific
- ⚠️ **PowerPoint Required** — Microsoft PowerPoint 2016 or later must be installed
- ⚠️ **Desktop Environment** — Controls actual PowerPoint process (not for server-side processing)

## 🎯 What You Can Do

**33 specialized tools with 204 operations:**

- 📄 **Slides** (1 tool, 8 ops) — Create, duplicate, move, delete, apply layouts, set name
- 🔷 **Shapes** (1 tool, 19 ops) — Add, move, resize, fill, line, shadow, rotation, z-order, grouping, copy, connectors, merge, flip, duplicate
- 📝 **Text** (1 tool, 5 ops) — Get/set text, find, replace, format
- 📊 **Charts** (1 tool, 5 ops) — Create charts, set title, type, get info, delete
- 📋 **Slide Tables** (1 tool, 8 ops) — Create, read, write cells, add/delete rows and columns, merge cells
- 🎬 **Animations** (1 tool, 4 ops) — List, add, remove, clear animation effects
- 🔄 **Transitions** (1 tool, 4 ops) — Get, set, remove, copy to all slides
- 🎨 **Design/Themes** (1 tool, 4 ops) — List designs, apply themes, get theme colors, list color schemes
- 🖼️ **Images** (1 tool, 1 op) — Insert images with position and size control
- 📝 **Notes** (1 tool, 4 ops) — Get, set, clear, append speaker notes
- 🏷️ **Sections** (1 tool, 4 ops) — List, add, rename, delete presentation sections
- 🔗 **Hyperlinks** (1 tool, 4 ops) — Add, read, list, remove hyperlinks
- 📺 **Slideshow** (1 tool, 4 ops) — Start, stop, navigate, get status
- 🎭 **Slide Masters** (1 tool, 1 op) — List masters and layouts
- 📤 **Export** (1 tool, 5 ops) — PDF, slide images, video (MP4), print, save-as (7 formats)
- 📝 **VBA** (1 tool, 5 ops) — List, view, import, delete, run macros
- 🎥 **Media** (1 tool, 3 ops) — Insert audio/video, get media info
- 🪟 **Window** (1 tool, 5 ops) — Get info, minimize, restore, maximize, set zoom
- 📁 **Files** (1 tool, 1 op) — File validation and info
- 📑 **Document Properties** (1 tool, 2 ops) — Get/set title, author, subject, etc.
- 💬 **Comments** (1 tool, 4 ops) — Add, list, delete, clear slide comments
- 📌 **Placeholders** (1 tool, 2 ops) — List placeholders, set placeholder text
- 🎨 **Slide Background** (1 tool, 4 ops) — Get info, set solid color, set image, reset to master
- 📋 **Headers & Footers** (1 tool, 2 ops) — Get/set footer text, slide numbers, date
- 🧩 **SmartArt** (1 tool, 2 ops) — Get diagram info, add nodes
- 📐 **Shape Alignment** (1 tool, 2 ops) — Align and distribute shapes on slides
- 🎪 **Custom Shows** (1 tool, 3 ops) — Create, list, delete custom slide shows
- 📏 **Page Setup** (1 tool, 2 ops) — Get/set slide size and orientation
- 📥 **Slide Import** (1 tool, 1 op) — Import slides from another .pptx file
- 🏷️ **Tags** (1 tool, 3 ops) — Custom metadata on slides and shapes

📚 **[Complete Feature Reference →](FEATURES.md)** — Detailed documentation of all 156 operations


## 💬 Example Prompts

**Create & Build Presentations:**
- *"Create a new PowerPoint presentation called QuarterlyReport.pptx with a title slide"*
- *"Add 5 slides with a 'Title and Content' layout"*
- *"Insert a company logo image on the first slide"*

**Content & Formatting:**
- *"Add a textbox on slide 2 with the text 'Q1 Revenue Summary' in bold 24pt Arial"*
- *"Create a table on slide 3 with columns for Region, Q1, Q2, Q3, Q4"*
- *"Set the shape fill color to #0078D4 and add a 2pt border"*

**Charts & Visuals:**
- *"Create a bar chart on slide 4 showing quarterly revenue data"*
- *"Set the chart title to 'Revenue by Quarter'"*
- *"Add an entrance animation to the chart shape"*

**Automation:**
- *"Export the presentation as PDF"*
- *"Run the FormatAllSlides macro"*
- *"Show me PowerPoint while you work"* — watch changes in real-time

**🪟 Agent Mode — Watch AI Work in PowerPoint:**
- *"Show me PowerPoint side-by-side while you build this presentation"* — real-time visibility
- *"Let me watch while you create the slides"*
- Status bar shows live progress: *"PptMcp: Creating chart on slide 4..."*

## 👥 Who Should Use This?

**Perfect for:**
- ✅ **Presenters** automating repetitive PowerPoint workflows
- ✅ **Developers** building PowerPoint-based reporting solutions
- ✅ **Business users** managing complex presentation decks
- ✅ **Teams** maintaining presentation templates and VBA macros

**Not suitable for:**
- ❌ Server-side processing (use libraries like Open XML SDK instead)
- ❌ Linux/macOS users (Windows + PowerPoint installation required)
- ❌ High-volume batch operations (consider PowerPoint-free alternatives)


## 🚀 Quick Start

| Platform | Installation |
|----------|-------------|
| **Any MCP Client** | `dotnet tool install --global PptMcp.McpServer` |
| **Details** | 📖 [Installation Guide](docs/INSTALLATION.md) |

**⚠️ Important:** Close all PowerPoint files before using. The server requires exclusive access to presentations during automation.


## 🔧 CLI vs MCP Server

This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`pptcli`) | Coding agents (Copilot, Cursor, Windsurf) | Fewer tokens — single tool, no large schemas. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. |

**Manual Installation:**
```powershell
# Install MCP Server and CLI
dotnet tool install --global PptMcp.McpServer
dotnet tool install --global PptMcp.CLI
```


## 🤖 Optional: Official Agent Client from Source

For larger deck-building tasks, this repo also ships an official source-side controller: `src\PptMcp.Agent`.

It is intentionally **not** a third server surface. Instead, it sits above the MCP server and runs one client-side loop:

- plan the deck
- execute through normal sequential MCP tool calls
- verify the generated deck
- repair incomplete output when needed

Quick start:

```powershell
dotnet build src\PptMcp.McpServer\PptMcp.McpServer.csproj -c Release

Set-Location src\PptMcp.Agent
npm install
npm run check
npm test

node .\src\cli.mjs run `
  --task "Build a 5-slide executive deck on Q4 revenue performance and next actions." `
  --output "C:\Users\you\Documents\q4-revenue-deck.pptx"
```

Read more:

- [Agent Client Component README](src/PptMcp.Agent/README.md)
- [Agent Client Architecture](docs/AGENT-CLIENT.md)
- [Eval Framework](eval/README.md)
- [Archetype Pipeline](docs/ARCHETYPE-PIPELINE.md)


## ⚙️ How It Works — COM Automation & Unified Service Architecture

**PptMcp uses Windows COM automation to control the actual PowerPoint application (not just .pptx files).**

Both the **MCP Server** and **CLI** communicate with a shared **PptMcp Service** that manages PowerPoint sessions. This unified architecture enables:

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server        │     │   CLI (pptcli)    │
│  (AI assistants)    │     │  (coding agents)    │
└─────────┬───────────┘     └─────────┬───────────┘
          │                           │
          └──────────┬────────────────┘
                     ▼
          ┌─────────────────────────┐
          │   PptMcp Service      │
          │  (shared session mgmt)  │
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   PowerPoint COM API    │
          │  (PowerPoint.Application)│
          └─────────────────────────┘
```

**Key Benefits:**
- ✅ **Shared Sessions** — CLI and MCP Server can access the same open presentations
- ✅ **Single PowerPoint Instance** — No duplicate processes or file locks
- ✅ **System Tray UI** — Monitor active sessions via the PptMcp tray icon

**💡 Tip: Watch PowerPoint While AI Works**
By default, PowerPoint runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me PowerPoint while you work"*
- *"Let me watch what you're doing"*
- *"Open PowerPoint so I can see the changes"*

The AI will display the PowerPoint window so you can watch every operation happen live!

## 📋 Additional Information

📚 **[CLI Guide →](src/PptMcp.CLI/README.md)** | **[MCP Server Guide →](src/PptMcp.McpServer/README.md)** | **[Agent Client →](src/PptMcp.Agent/README.md)** | **[Eval Framework →](eval/README.md)** | **[Archetype Pipeline →](docs/ARCHETYPE-PIPELINE.md)** | **[All Agent Skills →](skills/README.md)**

**License:** MIT License - see [LICENSE](LICENSE) file

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- **[Stefan Broenner (sbroenne)](https://github.com/sbroenne)** — Original author and creator of the [upstream mcp-server-ppt](https://github.com/sbroenne/mcp-server-ppt) project. This fork builds on his excellent foundation for PowerPoint COM automation via MCP.
- Microsoft PowerPoint Team — For comprehensive COM automation APIs
- Model Context Protocol community — For the AI integration standard
- Open Source Community — For inspiration and best practices

## Related Projects

Upstream projects by Stefan Broenner:

- [mcp-server-ppt (upstream)](https://github.com/sbroenne/mcp-server-ppt) — Original MCP Server for PowerPoint by Stefan Broenner
- [pytest-aitest](https://github.com/sbroenne/pytest-aitest) — LLM-powered testing framework for AI agents
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation
