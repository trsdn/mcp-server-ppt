# PptMcp.CLI - Command-Line Interface for PowerPoint Automation

[![NuGet](https://img.shields.io/nuget/v/PptMcp.CLI.svg)](https://www.nuget.org/packages/PptMcp.CLI)
[![Downloads](https://img.shields.io/nuget/dt/PptMcp.CLI.svg)](https://www.nuget.org/packages/PptMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Command-line interface for PowerPoint automation — preferred by coding agents.**

> **Published as its own .NET tool** - Install `PptMcp.CLI` to get the `pptcli` command. Install `PptMcp.McpServer` separately when you also need the MCP server (`mcp-ppt`).

The CLI provides 17 command categories with 225 operations matching the MCP Server. Uses **64% fewer tokens** than MCP Server because it wraps all operations in a single tool with skill-based guidance instead of loading 25 tool schemas into context.

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`pptcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection |

Also perfect for RPA workflows, CI/CD pipelines, batch processing, and automated testing.

➡️ **[Learn more and see examples](https://trsdn.github.io/mcp-server-ppt/)**

---

## 🚀 Quick Start

### Installation (.NET Global Tool - Recommended)

```powershell
# Install CLI tool
dotnet tool install --global PptMcp.CLI

# Verify installation
pptcli --version

# Get help
pptcli --help
```

> 🔁 **Session Workflow:** Always start with `pptcli session open <file>` (captures the session id), pass `--session <id>` to other commands, then `pptcli session close <id> --save` when finished. The CLI reuses the same PowerPoint instance through that lifecycle.

### Check for Updates

```powershell
# Check if newer version is available
pptcli version --check

# Update if available
dotnet tool update --global PptMcp.CLI
```

### Uninstall

```powershell
dotnet tool uninstall --global PptMcp.CLI
```

## 🤫 Quiet Mode (Agent-Friendly)

For scripting and coding agents, use `-q`/`--quiet` to suppress banner and output JSON only:

```powershell
pptcli -q session open data.pptx
pptcli -q range get-values --session 1 --sheet Sheet1 --range A1:B2
pptcli -q session close --session 1 --save
```

Banner auto-suppresses when stdout is piped or redirected.

## 🆘 Built-in Help

- `pptcli --help` – lists every command category plus the new descriptions from `Program.cs`
- `pptcli <command> --help` – shows verb-specific arguments (for example `pptcli sheet --help`)
- `pptcli session --help` – displays nested verbs such as `open`, `save`, `close`, and `list`

Descriptions are kept in sync with the CLI source so the help output always reflects the latest capabilities.

---

## ✨ Key Features

### 🔧 PowerPoint Development Automation
- **Power Query Management** - Export, import, update, and version control M code
- **VBA Development** - Manage VBA modules, run macros, automated testing
- **Data Model & DAX** - Create measures, manage relationships, Power Pivot operations
- **PivotTable Automation** - Create, configure, and manage PivotTables programmatically
- **Conditional Formatting** - Add rules (cell value, expression-based), clear formatting

### 📊 Data Operations
- **Slide Management** - Create, rename, copy, delete sheets with tab colors and visibility
- **Range Operations** - Read/write values, formulas, formatting, validation
- **PowerPoint Tables** - Lifecycle management, filtering, sorting, structured references
- **Connection Management** - OLEDB, ODBC, Text, Web connections with testing

### 🛡️ Production Ready
- **Zero Corruption Risk** - Uses PowerPoint's native COM API (not file manipulation)
- **Error Handling** - Comprehensive validation and helpful error messages
- **CI/CD Integration** - Perfect for automated workflows and testing
- **Windows Native** - Optimized for Windows PowerPoint automation

---

## 📋 Command Categories

PptMcp.CLI provides **225 operations** across 17 command categories:

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Full documentation with all operations

**Quick Reference:**

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 6 | `session create`, `session open` (IRM/AIP auto-detected), `session close`, `session list` |
| **Worksheets** | 16 | `sheet list`, `sheet create`, `sheet rename`, `sheet copy`, `sheet move`, `sheet copy-to-file` |
| **Power Query** | 10 | `powerquery list`, `powerquery create`, `powerquery refresh`, `powerquery update` |
| **Ranges** | 42 | `range get-values`, `range set-values`, `range copy`, `range find`, `range merge-cells` |
| **Conditional Formatting** | 2 | `conditionalformat add-rule`, `conditionalformat clear-rules` |
| **PowerPoint Tables** | 27 | `table create`, `table apply-filter`, `table get-data`, `table sort`, `table add-column` |
| **Charts** | 14 | `chart create-from-range`, `chart list`, `chart delete`, `chart move`, `chart fit-to-range` |
| **Chart Config** | 14 | `chartconfig set-title`, `chartconfig add-series`, `chartconfig set-style`, `chartconfig data-labels` |
| **PivotTables** | 30 | `pivottable create-from-range`, `pivottable add-row-field`, `pivottable refresh` |
| **Slicers** | 8 | `slicer create-slicer`, `slicer list-slicers`, `slicer set-slicer-selection` |
| **Data Model** | 19 | `datamodel create-measure`, `datamodel create-relationship`, `datamodel evaluate` |
| **Connections** | 9 | `connection list`, `connection refresh`, `connection test` |
| **Named Ranges** | 6 | `namedrange create`, `namedrange read`, `namedrange write`, `namedrange update` |
| **VBA** | 6 | `vba list`, `vba import`, `vba run`, `vba update` |
| **Calculation Mode** | 3 | `calculation get-mode`, `calculation set-mode`, `calculation calculate` |
| **Screenshot** | 2 | `screenshot capture`, `screenshot capture-sheet` |

**Note:** CLI uses session commands for multi-operation workflows.

---

## SESSION LIFECYCLE (Open/Save/Close)

The CLI uses an explicit session-based workflow where you open a file, perform operations, and optionally save before closing:

```powershell
# 1. Open a session
pptcli session open data.pptx
# Output: Session ID: 550e8400-e29b-41d4-a716-446655440000

# 2. List active sessions anytime
pptcli session list

# 3. Use the session ID with any commands
pptcli sheet create --session 550e8400-e29b-41d4-a716-446655440000 --sheet "NewSheet"
pptcli powerquery list --session 550e8400-e29b-41d4-a716-446655440000

# 4. Close and save changes
pptcli session close 550e8400-e29b-41d4-a716-446655440000 --save

# OR: Close and discard changes (no --save flag)
pptcli session close 550e8400-e29b-41d4-a716-446655440000
```

### Session Lifecycle Benefits

- **Explicit control** - Know exactly when changes are persisted with `--save`
- **Batch efficiency** - Keep single PowerPoint instance open for multiple operations (75-90% faster)
- **Flexibility** - Save and close in one command, or close without saving
- **Clean resource management** - Automatic PowerPoint cleanup when session closes

### Background Service & System Tray

When you run your first CLI command, the **PptMcp Service** starts automatically in the background. The service:

- **Manages PowerPoint COM** - Keeps PowerPoint instance alive between commands (no restart overhead)
- **Shows system tray icon** - Look for the PowerPoint icon in your Windows taskbar notification area
- **Tracks sessions** - Right-click the tray icon to see active sessions and close them
- **Shows session origin** - Sessions are labeled [CLI] or [MCP] showing which client created them
- **Auto-updates** - Notifies you when a new version is available and allows one-click updates

**Tray Icon Features:**
- 📋 **View sessions** - Double-click to see active session count
- 💾 **Close sessions** - Right-click → Sessions → select file → "Close Session..." (prompts to save with Cancel option)
- 🔄 **Update CLI** - When updates are available, click "Update to X.X.X" to update automatically
- ℹ️ **About** - Right-click → "About..." to see version info and helpful links
- 🛑 **Stop Service** - Right-click → "Stop Service" (prompts to save active sessions with Cancel option)

The service auto-stops after 10 minutes of inactivity (no active sessions).

---

## 💡 Command Reference

**Use `pptcli <command> --help` for complete parameter documentation.** The CLI help is always in sync with the code.

```powershell
pptcli --help              # List all commands
pptcli session --help      # Session lifecycle (open, close, save, list)
pptcli powerquery --help   # Power Query operations
pptcli range --help        # Cell/range operations
pptcli table --help        # PowerPoint Table operations
pptcli pivottable --help   # PivotTable operations
pptcli datamodel --help    # Data Model & DAX
pptcli vba --help          # VBA module management
```

### Typical Workflows

**Session-based automation (recommended):**
```powershell
pptcli -q session open report.pptx           # Returns session ID
pptcli -q sheet create --session 1 --sheet "Summary"
pptcli -q range set-values --session 1 --sheet Summary --range A1 --values '[["Hello"]]'
pptcli -q session close --session 1 --save   # Persist changes
```

**Power Query ETL:**
```powershell
pptcli powerquery create --session 1 --query "CleanData" --mcode-file transform.pq
pptcli powerquery refresh --session 1 --query "CleanData"
```

**PivotTable from Data Model:**
```powershell
pptcli pivottable create-from-datamodel --session 1 --table Sales --dest-sheet Analysis --dest-cell A1 --pivot-table SalesPivot
pptcli pivottable add-row-field --session 1 --pivot-table SalesPivot --field Region
pptcli pivottable add-value-field --session 1 --pivot-table SalesPivot --field Amount --function Sum
```

**VBA automation:**
```powershell
pptcli vba import --session 1 --module "Helpers" --code-file helpers.vba
pptcli vba run --session 1 --macro "Helpers.ProcessData"
```

---

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server 2016+ | COM interop is Windows-specific |
| **Microsoft PowerPoint** | PowerPoint 2016 or later | CLI controls actual PowerPoint application |
| **.NET 10 Runtime** | [Download](https://dotnet.microsoft.com/download/dotnet/10.0) | Required to run .NET global tools |

> **Note:** PptMcp.CLI controls the actual PowerPoint application via COM interop, not just file formats. This provides access to all PowerPoint features, but requires PowerPoint to be installed.

---

## 🔒 VBA Operations Setup (One-Time)

VBA commands require **"Trust access to the VBA project object model"** to be enabled:

1. Open PowerPoint
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice

This is a security setting that must be manually enabled. PptMcp.CLI never modifies security settings automatically.

For macro-enabled presentations, use `.pptm` extension:

```powershell
pptcli session create macros.pptm
# Returns session ID (e.g., 1)
pptcli vba import --session 1 --module MyModule --code-file code.vba
pptcli session close --session 1 --save
```

---

## 📖 Complete Documentation

- **[NuGet Package](https://www.nuget.org/packages/PptMcp.CLI)** - .NET Global Tool installation
- **[GitHub Repository](https://github.com/trsdn/mcp-server-ppt)** - Source code and issues
- **[Release Notes](https://github.com/trsdn/mcp-server-ppt/releases)** - Latest updates

---

## 🚧 Troubleshooting

### Command Not Found After Installation

```powershell
# Verify .NET tools path is in your PATH environment variable
dotnet tool list --global

# If pptcli is listed but not found, add .NET tools to PATH:
# The default location is: %USERPROFILE%\.dotnet\tools
```

### PowerPoint Not Found

```powershell
# Error: "Microsoft PowerPoint is not installed"
# Solution: Install Microsoft PowerPoint (any version 2016+)
```

### VBA Access Denied

```powershell
# Error: "Programmatic access to Visual Basic Project is not trusted"
# Solution: Enable VBA trust (see VBA Operations Setup above)
```

### Permission Issues

```powershell
# Run PowerShell/CMD as Administrator if you encounter permission errors
# Or install to user directory: dotnet tool install --global PptMcp.CLI
```

---

## 🛠️ Advanced Usage

### Scripting & Automation

```powershell
# PowerShell script example
$files = Get-ChildItem *.pptx
foreach ($file in $files) {
    $session = pptcli session open $file.Name | Select-String "Session ID: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    pptcli powerquery refresh --session $session --query "Sales Data"
    pptcli datamodel refresh --session $session
    pptcli session close $session --save
}
```

### CI/CD Integration

```yaml
# GitHub Actions example
- name: Install PptMcp.CLI
  run: dotnet tool install --global PptMcp.CLI

- name: Process PowerPoint Files
  run: |
    SESSION=$(pptcli session open data.pptx | grep "Session ID:" | cut -d' ' -f3)
    pptcli powerquery create --session $SESSION --query "Query1" --mcode-file queries/query1.pq
    pptcli powerquery refresh --session $SESSION --query "Query1"
    pptcli session close $SESSION --save
```


## ✅ Tested Scenarios

The CLI ships with real PowerPoint-backed integration tests that exercise the session lifecycle plus slide creation/listing flows through the same commands you run locally. Execute them with:

```powershell
dotnet test tests/PptMcp.CLI.Tests/PptMcp.CLI.Tests.csproj --filter "Layer=CLI"
```

These tests open actual presentations, issue `session open/list/close`, and call `pptcli sheet` actions to ensure the command pipeline stays healthy.

---

## 🤝 Related Tools

- **[PptMcp.McpServer](https://www.nuget.org/packages/PptMcp.McpServer)** - MCP server for AI assistant integration
- **[PowerPoint MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=trsdn.ppt-mcp)** - One-click PowerPoint automation in VS Code


---

## 📄 License

MIT License - see [LICENSE](../../LICENSE) for details.

---

## 🙋 Support

- **Issues**: [GitHub Issues](https://github.com/trsdn/mcp-server-ppt/issues)
- **Discussions**: [GitHub Discussions](https://github.com/trsdn/mcp-server-ppt/discussions)
- **Documentation**: [Complete Docs](../../docs/)

---

**Built with ❤️ for PowerPoint developers and automation engineers**
