# PowerPoint MCP Server - AI-Powered PowerPoint Automation

[![GitHub](https://img.shields.io/badge/GitHub-trsdn%2Fmcp--server--ppt-blue)](https://github.com/trsdn/mcp-server-ppt)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft PowerPoint with AI through GitHub Copilot - just ask in natural language!**

**MCP Server for PowerPoint** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate PowerPoint through natural language commands. Automate slide creation, layouts, shapes, text, charts, formatting, and transitions - no PowerPoint programming knowledge required. 

**🛡️ 100% Safe - Uses PowerPoint's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.pptx` files directly, this project uses PowerPoint's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in PowerPoint. Create a slide, run it, inspect the output, refine and repeat. PowerPoint becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real LLM workflows using [pytest-aitest](https://github.com/trsdn/pytest-aitest). We test that LLMs correctly understand and use our tools.

## Features

The PowerPoint MCP Server (ppt-mcp) provides **25 specialized tools with 225 operations** for comprehensive PowerPoint automation:

- 🔄 **Power Query** (1 tool, 11 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 18 ops) - Measures, relationships, model structure
- 🎨 **PowerPoint Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 26 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Slides** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-presentation moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management, presentation creation, IRM/AIP-protected file support
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for visual verification
- 🪧 **Window Management** (1 tool, 9 ops) - Show/hide PowerPoint, arrange, position, status bar feedback

📚 **[Complete Feature Reference →](https://github.com/trsdn/mcp-server-ppt/blob/main/FEATURES.md)**

### Agent Skills (Bundled)

This extension includes an **Agent Skill** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants:

- **[ppt-mcp](https://github.com/trsdn/mcp-server-ppt/blob/main/skills/ppt-mcp/SKILL.md)** - MCP Server tool guidance

**VS Code setup:** Enable the preview setting `chat.useAgentSkills` to allow Copilot to load skills. Skills are registered via VS Code's `chatSkills` contribution point and managed automatically.


## 💬 Example Prompts

**Create & Populate Data:**
- *"Create a new PowerPoint file called SalesTracker.pptx with slides for Date, Product, Quantity, Unit Price, and Total"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add sample data and a formula column for Quantity times Unit Price"*

**Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Import products.csv with Power Query, load to Data Model, create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*

**Formatting & Automation:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Export all Power Query M code to files for version control"*
- *"Show me PowerPoint while you work"* - watch changes in real-time


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all Power Query queries in presentation.pptx"
   - "Create a DAX measure for year-over-year revenue growth"
   - "Export all Power Queries and VBA modules to .vba files for version control"

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
