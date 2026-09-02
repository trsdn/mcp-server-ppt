# PptMcp - Model Context Protocol Server for PowerPoint

<!-- mcp-name: io.github.trsdn/mcp-server-ppt -->
mcp-name: io.github.trsdn/mcp-server-ppt

[![NuGet](https://img.shields.io/nuget/v/PptMcp.McpServer.svg)](https://www.nuget.org/packages/PptMcp.McpServer)
[![NuGet Downloads](https://img.shields.io/nuget/dt/PptMcp.McpServer.svg)](https://www.nuget.org/packages/PptMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/trsdn/mcp-server-ppt)

**Control PowerPoint with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered PowerPoint automation for slides, shapes, text, charts, and more.

➡️ **[Learn more and see examples](https://trsdn.github.io/mcp-server-ppt/)** 

**🛡️ 100% Safe - Uses PowerPoint's Native COM API**

Unlike third-party libraries that manipulate `.pptx` files (risking corruption), PptMcp uses **PowerPoint's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live PowerPoint files - see your changes happen in real-time.

**🔗 Unified Service Architecture** - The MCP Server forwards all requests to the shared PptMcp Service, enabling CLI and MCP to share sessions transparently.

**CLI also available:** The MCP Server tool (`mcp-ppt`) and CLI tool (`pptcli`) are published as separate .NET tools. Install `PptMcp.McpServer` for MCP clients, and optionally install `PptMcp.CLI` for scripting/RPA workflows.

**Requirements:** Windows OS + PowerPoint 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - install the `.vsix` from the [latest release](https://github.com/trsdn/mcp-server-ppt/releases/latest) for GitHub Copilot
2. **Manual Install** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
3. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/servers/io.github.trsdn/mcp-server-ppt)

**Manual Installation (All MCP Clients):**

Requires .NET 10 Runtime or SDK

```powershell
# Install MCP Server tool
dotnet tool install --global PptMcp.McpServer

# Optional: install CLI tool separately
dotnet tool install --global PptMcp.CLI
```

**Supported AI Assistants:**
- ✅ GitHub Copilot (VS Code, Visual Studio)
- ✅ Claude Desktop
- ✅ Cursor
- ✅ Cline (VS Code Extension)
- ✅ Windsurf
- ✅ Any MCP-compatible client

📖 **Detailed setup instructions:** [Installation Guide](https://github.com/trsdn/mcp-server-ppt/blob/main/docs/INSTALLATION.md)

🎯 **Quick config examples:** [examples/mcp-configs/](https://github.com/trsdn/mcp-server-ppt/tree/main/examples/mcp-configs)

## 🛠️ What You Can Do

**33 specialized tools with 224 operations:**

- 📁 **Files** (1 tool, 6 ops) - Session management, open/create/save/close presentations
- 📄 **Slides** (1 tool, 15 ops) - Lifecycle, layouts, hide/unhide, thumbnails, clone-with-replace
- 🔷 **Shapes** (1 tool, 35 ops) - Geometry, fill, gradient, line, shadow, glow, reflection, 3D, grouping, connectors, merge
- ✏️ **Text** (1 tool, 18 ops) - Get/set, find/replace, formatting, spacing, bullets, symbols, audits
- 📋 **Slide Tables** (1 tool, 13 ops) - Cells, rows, columns, merge, formatting, borders
- 📊 **Charts** (1 tool, 10 ops) - Create, data, title, type, legend, axes, data table
- 🖼️ **Images** (1 tool, 4 ops) - Insert, crop, brightness/contrast, transparent color
- 🎥 **Media** (1 tool, 4 ops) - Audio/video insertion, info, playback settings
- 🧩 **SmartArt** (1 tool, 6 ops) - Nodes, layout, style, level changes
- 🎨 **Design/Themes** (1 tool, 19 ops) - Themes, colors, fonts, archetypes, palettes, layout grids
- 🖌️ **Background** (1 tool, 5 ops) - Solid, gradient, image, reset to master
- 🎭 **Masters & Layouts** (1 tool, 5 ops) - Masters, layouts, master shape editing, cleanup
- 📌 **Placeholders** (1 tool, 3 ops) - List, set text, set image
- 📃 **Headers & Footers** (1 tool, 2 ops) - Footer text, slide numbers, date
- 📐 **Page Setup** (1 tool, 3 ops) - Slide size, orientation, first slide number
- ↔️ **Shape Alignment** (1 tool, 2 ops) - Align and distribute
- 🎬 **Animations** (1 tool, 6 ops) - Add, remove, timing, reorder
- 🔀 **Transitions** (1 tool, 4 ops) - Get, set, remove, apply to all
- 📺 **Slideshow** (1 tool, 5 ops) - Start, stop, navigate, status, configure
- 🎞️ **Custom Shows** (1 tool, 3 ops) - Create, list, delete
- 📂 **Sections** (1 tool, 4 ops) - List, add, rename, delete
- 📥 **Slide Import** (1 tool, 1 op) - Import slides from another presentation
- 📝 **Notes** (1 tool, 5 ops) - Speaker notes management
- 💬 **Comments** (1 tool, 4 ops) - Add, list, delete, clear
- 🔗 **Hyperlinks** (1 tool, 5 ops) - Add, get, list, remove, validate
- 🏷️ **Tags** (1 tool, 3 ops) - Custom metadata on slides and shapes
- 🗂️ **Document Properties** (1 tool, 4 ops) - Built-in and custom properties
- ♿ **Accessibility** (1 tool, 3 ops) - Audit and reading order
- 🔤 **Proofing** (1 tool, 3 ops) - Spell check and language
- 📤 **Export** (1 tool, 9 ops) - PDF, images, video, print, save-as, extract text/images
- 🖨️ **Print Options** (1 tool, 2 ops) - Print configuration
- ⚙️ **VBA** (1 tool, 5 ops) - Modules, import, execution
- 🪟 **Window Management** (1 tool, 7 ops) - Show/arrange PowerPoint, zoom, view mode

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Detailed documentation of all 224 operations

**AI-Powered Workflows:**
- 💬 Natural language PowerPoint commands through GitHub Copilot, Claude, or ChatGPT
- 🎨 Apply consistent themes, layouts, and formatting across a deck
- 📊 Build data-driven slides with charts, tables, and SmartArt
- 📋 Automate repetitive slide creation and cleanup
- 👀 **Show PowerPoint Mode** - Say "Show me PowerPoint while you work" to watch changes live


---

## 💡 Example Use Cases

**"Create a 10-slide investor pitch deck with a title slide, agenda, and a closing slide"**  
→ AI creates the presentation, applies a layout to each slide, and fills in placeholder content

**"Add a slide with a bar chart comparing quarterly sales by region"**  
→ AI adds the slide, creates the chart, sets its data and title, and positions it

**"Apply the corporate theme and add a fade transition to every slide"**  
→ AI applies the design, then sets the transition across the whole deck

**"Find every slide missing alt text and list the shapes"**  
→ AI runs the accessibility audit and reports shapes that need alt text

**"Export the deck to PDF and save each slide as a PNG"**  
→ AI exports the presentation and writes per-slide images to disk

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/trsdn/mcp-server-ppt)** - Source code, issues, discussions
- **[Installation Guide](https://github.com/trsdn/mcp-server-ppt/blob/main/docs/INSTALLATION.md)** - Detailed setup for all platforms
- **[VS Code Extension](https://github.com/trsdn/mcp-server-ppt/releases/latest)** - install the `.vsix` from the release assets
- **[CLI Documentation](https://github.com/trsdn/mcp-server-ppt/blob/main/src/PptMcp.CLI/README.md)** - Comprehensive commands for RPA and CI/CD automation

**License:** MIT  
**Platform:** Windows only (requires PowerPoint 2016+)  
**Support:** [GitHub Issues](https://github.com/trsdn/mcp-server-ppt/issues)
