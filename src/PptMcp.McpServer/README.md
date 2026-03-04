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

**CLI also available:** The MCP Server tool (`mcp-excel`) and CLI tool (`pptcli`) are published as separate .NET tools. Install `PptMcp.McpServer` for MCP clients, and optionally install `PptMcp.CLI` for scripting/RPA workflows.

**Requirements:** Windows OS + PowerPoint 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=trsdn.ppt-mcp) for GitHub Copilot
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

**25 specialized tools with 225 operations:**

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
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for visual verification
- 🪧 **Window Management** (1 tool, 9 ops) - Show/hide PowerPoint, arrange, position, status bar feedback

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Detailed documentation of all 225 operations

**AI-Powered Workflows:**
- 💬 Natural language PowerPoint commands through GitHub Copilot, Claude, or ChatGPT
- 🔄 Optimize Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show PowerPoint Mode** - Say "Show me PowerPoint while you work" to watch changes live


---

## 💡 Example Use Cases

**"Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns"**  
→ AI creates the presentation, adds headers, enters sample data, and builds formulas

**"Create a PivotTable from this data showing total sales by Product, then add a chart"**  
→ AI creates PivotTable, configures fields, and adds a linked visualization

**"Import products.csv with Power Query, load to Data Model, create a Total Revenue measure"**  
→ AI imports data, adds to Power Pivot, and creates DAX measures for analysis

**"Create a slicer for the Region field so I can filter interactively"**  
→ AI adds slicers connected to PivotTables or Tables for point-and-click filtering

**"Put this data in A1: Name, Age / Alice, 30 / Bob, 25"**  
→ AI writes data directly to cells using natural delimiters you provide

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/trsdn/mcp-server-ppt)** - Source code, issues, discussions
- **[Installation Guide](https://github.com/trsdn/mcp-server-ppt/blob/main/docs/INSTALLATION.md)** - Detailed setup for all platforms
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=trsdn.ppt-mcp)** - One-click installation
- **[CLI Documentation](https://github.com/trsdn/mcp-server-ppt/blob/main/src/PptMcp.CLI/README.md)** - Comprehensive commands for RPA and CI/CD automation

**License:** MIT  
**Privacy:** [PRIVACY.md](https://github.com/trsdn/mcp-server-ppt/blob/main/PRIVACY.md)  
**Platform:** Windows only (requires PowerPoint 2016+)  
**Support:** [GitHub Issues](https://github.com/trsdn/mcp-server-ppt/issues)
