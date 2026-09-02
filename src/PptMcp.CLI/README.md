# PptMcp.CLI - Command-Line Interface for PowerPoint Automation

[![NuGet](https://img.shields.io/nuget/v/PptMcp.CLI.svg)](https://www.nuget.org/packages/PptMcp.CLI)
[![Downloads](https://img.shields.io/nuget/dt/PptMcp.CLI.svg)](https://www.nuget.org/packages/PptMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Command-line interface for PowerPoint automation — preferred by coding agents.**

> **Published as its own .NET tool** - Install `PptMcp.CLI` to get the `pptcli` command. Install `PptMcp.McpServer` separately when you also need the MCP server (`mcp-ppt`).

The CLI provides 33 command categories with 224 operations matching the MCP Server. Uses **64% fewer tokens** than MCP Server because it wraps all operations in a single tool with skill-based guidance instead of loading 33 tool schemas into context.

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
pptcli -q shape list --session 1 --slide-index 1
pptcli -q session close --session 1 --save
```

Banner auto-suppresses when stdout is piped or redirected.

## 🆘 Built-in Help

- `pptcli --help` – lists every command category plus the new descriptions from `Program.cs`
- `pptcli <command> --help` – shows verb-specific arguments (for example `pptcli slide --help`)
- `pptcli session --help` – displays nested verbs such as `open`, `save`, `close`, `list`, and `test`
- `pptcli list-actions [command]` – JSON list of valid actions, for agents that would rather parse than scrape `--help`

Descriptions are kept in sync with the CLI source so the help output always reflects the latest capabilities.

---

## ✨ Key Features

### 🔧 PowerPoint Deck Automation
- **Slide Management** - Create, duplicate, move, delete, apply layouts, hide/unhide
- **Shape Manipulation** - Geometry, fill, gradient, line, shadow, glow, 3D, grouping, connectors
- **Text & Typography** - Get/set text, find/replace, formatting, spacing, bullets, symbols
- **VBA Development** - Manage VBA modules, run macros, automated testing
- **Design Systems** - Themes, palettes, archetypes, layout grids, style profiles

### 📊 Content Operations
- **Slide Tables** - Cells, rows, columns, merging, formatting, borders
- **Charts** - Create, set data, titles, types, legends, axes
- **Images & Media** - Insert, crop, adjust, audio/video with playback settings
- **SmartArt** - Nodes, layouts, styles, level changes

### 🛡️ Production Ready
- **Zero Corruption Risk** - Uses PowerPoint's native COM API (not file manipulation)
- **Error Handling** - Comprehensive validation and helpful error messages
- **CI/CD Integration** - Perfect for automated workflows and testing
- **Windows Native** - Optimized for Windows PowerPoint automation

---

## 📋 Command Categories

PptMcp.CLI provides **224 operations** across 33 command categories:

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Full documentation with all operations

**Quick Reference:**

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 6 | `session create`, `session open` (IRM/AIP auto-detected), `session close --save`, `session list`, `session test` |
| **Slide** | 15 | `slide list`, `slide create`, `slide duplicate`, `slide apply-layout` |
| **Shape** | 35 | `shape add-shape`, `shape move-resize`, `shape set-fill`, `shape group` |
| **Text** | 18 | `text set`, `text replace`, `text format`, `text set-bullets` |
| **Slide Table** | 13 | `slidetable create`, `slidetable write-cell`, `slidetable merge-cells` |
| **Chart** | 10 | `chart create`, `chart set-data`, `chart set-title`, `chart set-legend` |
| **Design** | 19 | `design apply-theme`, `design get-colors`, `design list-palettes` |
| **Export** | 9 | `export to-pdf`, `export slide-to-image`, `export to-video` |
| **Window** | 7 | `window get-info`, `window maximize`, `window set-zoom` |
| **Animation** | 6 | `animation add`, `animation set-timing`, `animation reorder` |
| **SmartArt** | 6 | `smartart add-node`, `smartart set-layout`, `smartart set-style` |
| **Background** | 5 | `background set-color`, `background set-image`, `background reset` |
| **Master & Layout** | 5 | `master list`, `master list-layouts`, `master edit-shape-text` |
| **Notes** | 5 | `notes get`, `notes set`, `notes append`, `notes read-all` |
| **Hyperlink** | 5 | `hyperlink add`, `hyperlink list`, `hyperlink validate` |
| **Slideshow** | 5 | `slideshow start`, `slideshow goto-slide`, `slideshow configure` |
| **VBA** | 5 | `vba list`, `vba import`, `vba run`, `vba delete` |
| **Image** | 4 | `image insert`, `image crop`, `image set-brightness-contrast` |
| **Media** | 4 | `media insert-audio`, `media insert-video`, `media set-playback` |
| **Comment** | 4 | `comment add`, `comment list`, `comment delete` |
| **Section** | 4 | `section add`, `section rename`, `section delete` |
| **Transition** | 4 | `transition set`, `transition remove`, `transition copy-to-all` |
| **Document Property** | 4 | `docproperty get`, `docproperty set`, `docproperty set-custom` |
| **Accessibility** | 3 | `accessibility audit`, `accessibility set-reading-order` |
| **Proofing** | 3 | `proofing check-spelling`, `proofing set-language` |
| **Placeholder** | 3 | `placeholder list`, `placeholder set-text`, `placeholder set-image` |
| **Page Setup** | 3 | `pagesetup set-size`, `pagesetup set-first-number` |
| **Custom Show** | 3 | `customshow create`, `customshow list`, `customshow delete` |
| **Tag** | 3 | `tag set`, `tag get`, `tag delete` |
| **Header & Footer** | 2 | `headerfooter get`, `headerfooter set` |
| **Shape Alignment** | 2 | `shapealign align`, `shapealign distribute` |
| **Print Options** | 2 | `printoptions get`, `printoptions set` |
| **Slide Import** | 1 | `slideimport import` |

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
pptcli slide list --session 550e8400-e29b-41d4-a716-446655440000
pptcli shape list --session 550e8400-e29b-41d4-a716-446655440000 --slide-index 1

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
pptcli session --help      # Session lifecycle (create, open, close, save, list)
pptcli slide --help        # Slide operations
pptcli shape --help        # Shape operations
pptcli text --help         # Text operations
pptcli chart --help        # Chart operations
pptcli design --help       # Themes and design systems
pptcli export --help       # PDF, images, video export
pptcli vba --help          # VBA module management
```

### Typical Workflows

**Session-based automation (recommended):**
```powershell
pptcli -q session open report.pptx           # Returns session ID
pptcli -q slide create --session 1 --layout-name "Title and Content"
pptcli -q text set --session 1 --slide-index 1 --shape-name "Title 1" --text "Q1 Review"
pptcli -q session close --session 1 --save   # Persist changes
```

**Build a chart slide:**
```powershell
pptcli -q slide create --session 1
pptcli -q chart create --session 1 --slide-index 2 --chart-type ColumnClustered
pptcli -q chart set-title --session 1 --slide-index 2 --shape-name "Chart 1" --title "Sales by Region"
```

**Apply design and transitions:**
```powershell
pptcli -q design list --session 1
pptcli -q transition set --session 1 --slide-index 1 --transition-type Fade
pptcli -q transition copy-to-all --session 1 --slide-index 1
```

**Export a finished deck:**
```powershell
pptcli -q export to-pdf --session 1 --destination-path C:\out\deck.pdf
pptcli -q export all-slides-to-images --session 1 --destination-directory C:\out\slides
```

**VBA automation:**
```powershell
pptcli vba import --session 1 --module-name "Helpers" --code "Sub ProcessData()`nEnd Sub"
pptcli vba run --session 1 --macro-name "Helpers.ProcessData"
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
    pptcli slide list --session $session
    pptcli export pdf --session $session --destination-path "$($file.BaseName).pdf"
    pptcli session close $session --save
}
```

### CI/CD Integration

```yaml
# GitHub Actions example
- name: Install PptMcp.CLI
  run: dotnet tool install --global PptMcp.CLI

- name: Process PowerPoint Files
  shell: pwsh
  run: |
    $session = pptcli session open deck.pptx | Select-String "Session ID: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    pptcli slide list --session $session
    pptcli export pdf --session $session --destination-path deck.pdf
    pptcli session close $session --save
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
- **[PowerPoint MCP VS Code Extension](https://github.com/trsdn/mcp-server-ppt/releases/latest)** - PowerPoint automation in VS Code (install the `.vsix` from the release)


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
