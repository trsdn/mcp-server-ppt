# PowerPoint MCP Server - AI-Powered PowerPoint Automation

[![GitHub](https://img.shields.io/badge/GitHub-trsdn%2Fmcp--server--ppt-blue)](https://github.com/trsdn/mcp-server-ppt)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft PowerPoint with AI through GitHub Copilot - just ask in natural language!**

**MCP Server for PowerPoint** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate PowerPoint through natural language commands. Automate slide creation, layouts, shapes, text, charts, formatting, and transitions - no PowerPoint programming knowledge required. 

**🛡️ 100% Safe - Uses PowerPoint's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.pptx` files directly, this project uses PowerPoint's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in PowerPoint. Create a slide, run it, inspect the output, refine and repeat. PowerPoint becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real LLM workflows using [pytest-aitest](https://github.com/trsdn/pytest-aitest). We test that LLMs correctly understand and use our tools.

## Features

The PowerPoint MCP Server (ppt-mcp) provides **33 specialized tools with 223 operations** for comprehensive PowerPoint automation:

- 📁 **Files** (1 tool, 6 ops) - Session management, presentation creation, IRM/AIP-protected file support
- 📄 **Slides** (1 tool, 15 ops) - Lifecycle, layouts, hide/unhide, thumbnails
- 🔷 **Shapes** (1 tool, 35 ops) - Geometry, fill, line, effects, grouping, connectors, merge
- ✏️ **Text** (1 tool, 18 ops) - Get/set, find/replace, formatting, spacing, bullets
- 📋 **Slide Tables** (1 tool, 13 ops) - Cells, rows, columns, merge, formatting
- 📊 **Charts** (1 tool, 10 ops) - Create, data, title, type, legend, axes
- 🎨 **Design/Themes** (1 tool, 19 ops) - Themes, colors, fonts, archetypes, palettes
- 🎭 **Masters & Layouts** (1 tool, 5 ops) - Masters, layouts, master shape editing
- 🎬 **Animations & Transitions** (2 tools, 10 ops) - Effects, timing, reorder, deck-wide transitions
- 🖼️ **Images & Media** (2 tools, 8 ops) - Insert, crop, adjust, audio/video
- 🧩 **SmartArt** (1 tool, 6 ops) - Nodes, layout, style, level changes
- 📤 **Export** (1 tool, 9 ops) - PDF, images, video, print, extract text/images
- ♿ **Accessibility & Proofing** (2 tools, 6 ops) - Alt-text audit, reading order, spell check
- ⚙️ **VBA** (1 tool, 5 ops) - Modules, import, execution
- 🪧 **Window Management** (1 tool, 7 ops) - Show/hide PowerPoint, arrange, zoom, view mode

📚 **[Complete Feature Reference →](https://github.com/trsdn/mcp-server-ppt/blob/main/FEATURES.md)**

### Agent Skills (Bundled)

This extension includes an **Agent Skill** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants:

- **[ppt-mcp](https://github.com/trsdn/mcp-server-ppt/blob/main/skills/ppt-mcp/SKILL.md)** - MCP Server tool guidance

**VS Code setup:** Enable the preview setting `chat.useAgentSkills` to allow Copilot to load skills. Skills are registered via VS Code's `chatSkills` contribution point and managed automatically.


## 💬 Example Prompts

**Create & Build Decks:**
- *"Create a new PowerPoint file called SalesReview.pptx with a title slide and three content slides"*
- *"Add a slide with the title 'Q1 Results' and three bullet points summarizing the quarter"*
- *"Duplicate slide 3 and replace the product name with 'Contoso Pro'"*

**Visualization & Layout:**
- *"Add a bar chart to slide 4 comparing total sales by region"*
- *"Insert a table with product sales data on a new slide and merge the header cells"*
- *"Add a SmartArt process diagram with four steps to slide 2"*

**Formatting & Automation:**
- *"Apply the corporate theme and add a fade transition to every slide"*
- *"Audit the deck for missing alt text and fix the shapes you find"*
- *"Show me PowerPoint while you work"* - watch changes in real-time


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all slides in presentation.pptx"
   - "Add a closing slide with a thank-you message"
   - "Export the deck to PDF and save each slide as a PNG"

**That's it!** The extension includes a self-contained MCP server - no .NET runtime or SDK needed.

➡️ **[Learn more and see examples](https://trsdn.github.io/mcp-server-ppt/)**

## Requirements

- **Windows OS** - PowerPoint COM automation requires Windows
- **Microsoft PowerPoint 2016+** - Must be installed on your system

## Potential Issues

**"PowerPoint is not installed" error:**
- Ensure Microsoft PowerPoint 2016+ is installed on your Windows machine
- Try opening PowerPoint manually to verify it works

**"VBA access denied" error:**
- VBA operations require one-time manual setup in PowerPoint
- Go to: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Check "Trust access to the VBA project object model"

**Copilot doesn't see PowerPoint tools:**
- Restart VS Code after installing the extension
- ### Troubleshooting

- Check Output panel → "PowerPoint MCP Server" for connection status

## Documentation & Support

- **[Complete Documentation](https://github.com/trsdn/mcp-server-ppt)** - Full guides and examples
- **[Report Issues](https://github.com/trsdn/mcp-server-ppt/issues)** - Bug reports and feature requests

## License & Privacy

MIT License - see [LICENSE](https://github.com/trsdn/mcp-server-ppt/blob/main/LICENSE)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
