# PptMcp - MCP Server for Microsoft PowerPoint

[![Release](https://img.shields.io/github/v/release/trsdn/mcp-server-ppt?display_name=tag)](https://github.com/trsdn/mcp-server-ppt/releases)
[![Release Workflow](https://github.com/trsdn/mcp-server-ppt/actions/workflows/release.yml/badge.svg)](https://github.com/trsdn/mcp-server-ppt/actions/workflows/release.yml)
[![NuGet MCP Server](https://img.shields.io/nuget/v/PptMcp.McpServer?label=NuGet%20MCP%20Server)](https://www.nuget.org/packages/PptMcp.McpServer)
[![NuGet CLI](https://img.shields.io/nuget/v/PptMcp.CLI?label=NuGet%20CLI)](https://www.nuget.org/packages/PptMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9-blue.svg)](https://dotnet.microsoft.com/download/dotnet/9.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/trsdn/mcp-server-ppt)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

[![Install in VS Code](https://img.shields.io/badge/VS_Code-Install_Server-0098FF?style=flat-square)](https://insiders.vscode.dev/redirect?url=vscode%3Amcp%2Finstall%3F%257B%2522name%2522%253A%2522ppt-mcp%2522%252C%2522command%2522%253A%2522mcp-ppt%2522%257D)
[![Install in VS Code Insiders](https://img.shields.io/badge/VS_Code_Insiders-Install_Server-24bfa5?style=flat-square)](https://insiders.vscode.dev/redirect?url=vscode-insiders%3Amcp%2Finstall%3F%257B%2522name%2522%253A%2522ppt-mcp%2522%252C%2522command%2522%253A%2522mcp-ppt%2522%257D)
[![Install in Cursor](https://img.shields.io/badge/Cursor-Install_Server-000000?style=flat-square)](https://cursor.com/en/install-mcp?name=ppt-mcp&config=eyJjb21tYW5kIjoibWNwLXBwdCJ9)

> The install badges register the server as `ppt-mcp` running the `mcp-ppt` command. Install the tool
> first with `dotnet tool install --global PptMcp.McpServer`, otherwise the client has nothing to launch.

**Automate PowerPoint with AI — A Model Context Protocol (MCP) server for comprehensive PowerPoint automation through conversational AI.**

**PptMcp** enables AI assistants such as GitHub Copilot, Claude, and ChatGPT to automate Microsoft PowerPoint through natural language commands. It covers real presentation work end to end: create slides, edit text, place shapes and images, build charts and tables, apply themes, run VBA, export to PDF or video, and manage live PowerPoint windows safely through the native COM API.

**Origin and credit:** This project builds on the original MCP automation foundation created by **[Stefan Broenner (sbroenne)](https://github.com/sbroenne)** in [mcp-server-excel](https://github.com/sbroenne/mcp-server-excel), extended here into a full PowerPoint-focused toolchain.

**Releases and packages:** [GitHub Releases](https://github.com/trsdn/mcp-server-ppt/releases) | [NuGet: PptMcp.McpServer](https://www.nuget.org/packages/PptMcp.McpServer) | [NuGet: PptMcp.CLI](https://www.nuget.org/packages/PptMcp.CLI)

For multi-phase build / verify / repair workflows from source, the repo also includes the official orchestration client under `src\PptMcp.Agent`.

**🛡️ 100% Safe — Uses PowerPoint's Native COM API** — Zero risk of file corruption. Uses PowerPoint's official COM API ensuring complete safety and compatibility.

**💡 Interactive Development** — See results instantly in PowerPoint. Add slides, create charts, format text, and iterate. PowerPoint becomes your AI-powered workspace.

**Technical Requirements:**
- ⚠️ **Windows Only** — COM interop is Windows-specific
- ⚠️ **PowerPoint Required** — Microsoft PowerPoint 2016 or later must be installed
- ⚠️ **Desktop Environment** — Controls actual PowerPoint process (not for server-side processing)

## 🎯 What You Can Do

**33 specialized tools with 224 operations:**

- 📁 **Files** (1 tool, 6 ops) — Open, close, create, save, list, validate presentations
- 📄 **Slides** (1 tool, 15 ops) — Create, duplicate, move, delete, apply layouts, hide/unhide, thumbnails, clone-with-replace
- 🔷 **Shapes** (1 tool, 35 ops) — Add, move, resize, fill, gradient, line, shadow, glow, reflection, opacity, 3D, rotation, z-order, grouping, connectors, merge, flip, scale
- ✏️ **Text** (1 tool, 18 ops) — Get/set text, find, replace, format, spacing, bullets, case, symbols, alt-text audit
- 📋 **Slide Tables** (1 tool, 13 ops) — Create, read/write cells and rows, add/delete rows and columns, merge, format, borders
- 📊 **Charts** (1 tool, 10 ops) — Create, set data, title, type, legend, axis titles, data table
- 🖼️ **Images** (1 tool, 4 ops) — Insert, crop, brightness/contrast, transparent color
- 🎥 **Media** (1 tool, 4 ops) — Insert audio/video, media info, playback settings
- 🧩 **SmartArt** (1 tool, 6 ops) — Diagram info, add/delete nodes, layout, style, level changes
- 🎨 **Design/Themes** (1 tool, 19 ops) — Themes, color schemes, fonts, archetypes, palettes, layout grids, style profiles
- 🖌️ **Slide Background** (1 tool, 5 ops) — Get, solid color, gradient, image, reset to master
- 🎭 **Masters & Layouts** (1 tool, 5 ops) — List masters and layouts, edit master shapes, delete unused
- 📌 **Placeholders** (1 tool, 3 ops) — List placeholders, set text, set image
- 📃 **Headers & Footers** (1 tool, 2 ops) — Get/set footer text, slide numbers, date
- 📐 **Page Setup** (1 tool, 3 ops) — Slide size, orientation, first slide number
- ↔️ **Shape Alignment** (1 tool, 2 ops) — Align and distribute shapes on slides
- 🎬 **Animations** (1 tool, 6 ops) — List, add, remove, clear, timing, reorder
- 🔀 **Transitions** (1 tool, 4 ops) — Get, set, remove, copy to all slides
- 📺 **Slideshow** (1 tool, 5 ops) — Start, stop, navigate, status, configure
- 🎞️ **Custom Shows** (1 tool, 3 ops) — Create, list, delete custom slide shows
- 📂 **Sections** (1 tool, 4 ops) — List, add, rename, delete presentation sections
- 📥 **Slide Import** (1 tool, 1 op) — Import slides from another .pptx file
- 📝 **Notes** (1 tool, 5 ops) — Get, set, clear, append, read all speaker notes
- 💬 **Comments** (1 tool, 4 ops) — Add, list, delete, clear slide comments
- 🔗 **Hyperlinks** (1 tool, 5 ops) — Add, get, list, remove, validate hyperlinks
- 🏷️ **Tags** (1 tool, 3 ops) — Custom metadata on slides and shapes
- 🗂️ **Document Properties** (1 tool, 4 ops) — Get/set built-in and custom properties
- ♿ **Accessibility** (1 tool, 3 ops) — Audit, get/set reading order
- 🔤 **Proofing** (1 tool, 3 ops) — Spell check, get/set language
- 📤 **Export** (1 tool, 9 ops) — PDF, slide images, video (MP4), print, save-as, extract text/images
- 🖨️ **Print Options** (1 tool, 2 ops) — Get/set print configuration
- ⚙️ **VBA** (1 tool, 5 ops) — List, view, import, delete, run macros
- 🪟 **Window** (1 tool, 7 ops) — Info, minimize, restore, maximize, zoom, view mode

📚 **[Complete Feature Reference →](FEATURES.md)** — Detailed documentation of all 224 operations


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
- Microsoft PowerPoint Team — For comprehensive COM automation APIs
- Model Context Protocol community — For the AI integration standard
- Open Source Community — For inspiration and best practices

## Related Projects

Upstream projects by Stefan Broenner:

- [mcp-server-excel (upstream)](https://github.com/sbroenne/mcp-server-excel) — Original MCP Server for Excel by Stefan Broenner
- [pytest-aitest](https://github.com/sbroenne/pytest-aitest) — LLM-powered testing framework for AI agents
