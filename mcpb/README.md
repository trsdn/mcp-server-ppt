# PowerPoint (Windows)

**Automate Microsoft PowerPoint with Claude** - Control PowerPoint through natural language conversations. Requires Windows and local Office install.

## What It Does

PowerPoint MCP Server lets you automate PowerPoint through conversation with Claude:

- **Create & Edit** - Build presentations, slides, and shapes
- **Analyze Data** - Charts, tables, and SmartArt
- **Transform Data** - Slide animations and transitions
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

### Example 1: Create a Sales Presentation

**You say:** *"Create a new PowerPoint file called SalesPresentation.pptx with a title slide, a slide with bullet points summarizing Q1 results, and a slide with a bar chart comparing sales by region."*

**What happens:**
- Creates a new presentation
- Adds a title slide with the presentation name
- Creates a content slide with Q1 summary bullet points
- Adds a chart slide with a bar chart comparing regional sales
- Applies consistent formatting and theme
- Confirms completion with file location

### Example 2: Build a Dashboard Slide

**You say:** *"Add a new slide with a table showing product sales data and a pie chart next to it visualizing the breakdown by category."*

**What happens:**
- Creates a new slide with a split layout
- Adds a table with product sales data
- Creates a pie chart showing category breakdown
- Positions both elements side by side
- Returns confirmation with slide number

### Example 3: Create Professional Presentation

**You say:** *"Create a 10-slide investor pitch deck with a title slide, agenda, market overview with a chart, product features with SmartArt, team bios, and a closing slide. Use a professional blue theme with slide transitions."*

**What happens:**
- Creates a new presentation with a professional blue theme
- Builds all 10 slides with appropriate layouts
- Adds charts and SmartArt graphics where specified
- Applies consistent formatting and slide transitions
- Confirms the presentation is ready for review

---

**More things you can ask:**

- *"Show me PowerPoint side-by-side while you build this presentation"* - Agent Mode: watch every step happen live
- *"Add a slide with a table showing Name, Role, and Department for the team"*
- *"Add slide transitions and entrance animations to all slides"*
- *"Apply the company brand colors and format the title slides consistently"*
- *"Create a SmartArt diagram showing our organizational structure"*
- *"Run the UpdateSlides macro"*
- *"Show me PowerPoint while you work"* - watch changes in real-time

## Tips for Best Results

- **Be specific** - Include file paths, slide numbers, and shape references when you know them
- **Start simple** - Build complex presentations step by step
- **Ask to see PowerPoint** - Say *"Show me PowerPoint while you work"* to watch changes in real-time
- **Close files first** - PowerPoint MCP needs exclusive access to presentations during automation

## Privacy & Security

PowerPoint MCP Server runs **entirely on your computer**. Your PowerPoint data:
- Never leaves your machine
- Is not sent to any external servers
- Is not used for training AI models

**Zero Logging:** This software does not collect any telemetry, usage statistics, or analytics data. No data is transmitted to external services.

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
