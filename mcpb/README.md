# PowerPoint (Windows)

**Automate Microsoft PowerPoint with Claude** - Control PowerPoint through natural language conversations. Requires Windows and local Office install.

## What It Does

PowerPoint MCP Server lets you automate PowerPoint through conversation with Claude:

- **Create & Edit** - Build presentations, slides, and shapes
- **Analyze Data** - Charts, tables, and SmartArt
- **Transform Data** - Power Query imports and transformations
- **Format & Style** - Themes, layouts, transitions, animations
- **Automate** - VBA macros, batch operations
- **Agent Mode** - Say "show me PowerPoint" and watch AI work in real-time, side-by-side with Claude

**25 tools with 225 operations** for comprehensive PowerPoint automation.

## Requirements

- **Windows** (required - uses PowerPoint COM automation)
- **Microsoft PowerPoint 2016 or later**
- **Claude Desktop** (Windows version)

## Installation

1. Download the `.mcpb` file from the [latest release](https://github.com/trsdn/mcp-server-ppt/releases/latest)
2. Double-click to install in Claude Desktop
3. Restart Claude Desktop if prompted

That's it! Start a new conversation and ask Claude to work with PowerPoint.

## Usage Examples

These examples work with any PowerPoint file, including a new empty presentation.

### Example 1: Create a Sales Tracker

**You say:** *"Create a new PowerPoint file called SalesTracker.pptx with slides for tracking sales. Include columns for Date, Product, Quantity, Unit Price, and Total. Add some sample data and a formula for the Total column."*

**What happens:**
- Creates a new presentation
- Adds column headers (Date, Product, Quantity, Unit Price, Total)
- Enters sample sales data
- Creates formulas in the Total column (Quantity × Unit Price)
- Formats the data as a PowerPoint Table
- Confirms completion with file location

### Example 2: Build a Dashboard with PivotTable and Chart

**You say:** *"I want to analyze this data. Create a PivotTable that shows total sales by Product, then add a bar chart to visualize the results."*

**What happens:**
- Creates a PivotTable from the data
- Configures Product as rows and Total as sum values
- Creates a new slide for the PivotTable
- Adds a bar chart based on the PivotTable
- Returns confirmation with locations of both

### Example 3: Power Query and Data Model Analysis

**You say:** *"Use Power Query to import this CSV file: C:/Data/products.csv. Add the data to the Data Model and create measures for Total Revenue and Average Rating."*

**What happens:**
- Imports the CSV using Power Query
- Loads the data to a slide as a PowerPoint Table
- Adds the table to the Power Pivot Data Model
- Creates DAX measures for analysis
- Confirms the data is ready for PivotTable analysis

---

**More things you can ask:**

- *"Show me PowerPoint side-by-side while you build this dashboard"* - Agent Mode: watch every step happen live
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Create a relationship between the Orders and Products tables using ProductID"*
- *"Run the UpdatePrices macro"*
- *"Show me PowerPoint while you work"* - watch changes in real-time

## Tips for Best Results

- **Be specific** - Include file paths, sheet names, and column references when you know them
- **Start simple** - Build complex presentations step by step
- **Ask to see PowerPoint** - Say *"Show me PowerPoint while you work"* to watch changes in real-time
- **Close files first** - PowerPoint MCP needs exclusive access to presentations during automation

## Privacy & Security

PowerPoint MCP Server runs **entirely on your computer**. Your PowerPoint data:
- Never leaves your machine
- Is not sent to any external servers
- Is not used for training AI models

**Anonymous Telemetry:** We collect anonymous usage statistics (tool usage, performance metrics, error rates) to improve the software. No file contents, file names, or personal data are collected.

See our complete [Privacy Policy](https://PptMcpserver.dev/privacy/).

## Troubleshooting

**Claude says the tool isn't available:**
- Restart Claude Desktop after installation
- Check Settings → Integrations to verify PowerPoint MCP Server is enabled

**PowerPoint operations fail:**
- Close the presentation in PowerPoint before asking Claude to modify it
- Ensure PowerPoint is installed and working normally

**Need help?**
- [Report an issue](https://github.com/trsdn/mcp-server-ppt/issues)
- [Full documentation](https://PptMcpserver.dev/)

## Links

- [GitHub Repository](https://github.com/trsdn/mcp-server-ppt)
- [Feature Reference](https://PptMcpserver.dev/features/)
- [Agent Skills](https://github.com/trsdn/mcp-server-ppt/blob/main/skills/README.md) - Cross-platform AI guidance
- [Privacy Policy](https://PptMcpserver.dev/privacy/)
- [License (MIT)](https://github.com/trsdn/mcp-server-ppt/blob/main/LICENSE)
