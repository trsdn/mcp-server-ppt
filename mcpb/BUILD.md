# MCPB Build & Packaging Guide

This document contains developer information for building and submitting the PowerPoint MCP Server to the Claude Desktop directory.

## Directory Contents

```
mcpb/
├── Build-McpBundle.ps1   # Packaging script
├── manifest.json         # MCPB manifest for Claude directory
├── icon-512.png          # Server icon (512x512 PNG)
├── README.md             # End-user documentation (ships with package)
├── BUILD.md              # This file (developer documentation)
└── artifacts/            # Build output (gitignored)
```

## Prerequisites

- .NET 10 SDK
- Windows x64 development machine

## Building the MCPB Package

From the `mcpb` directory:

```powershell
.\Build-McpBundle.ps1
```

This creates `mcpb/artifacts/PptMcp.McpServer-win-x64.zip`.

### Build Options

```powershell
# Specify version
.\Build-McpBundle.ps1 -Version "1.2.0"

# Custom output directory
.\Build-McpBundle.ps1 -OutputDir "./dist"
```

## Package Contents

The MCPB zip file contains:

```
PptMcp.McpServer-win-x64.zip/
├── PptMcp.McpServer.exe  # Self-contained executable (~15 MB)
├── .mcp/
│   └── server.json                   # MCP server configuration
├── manifest.json                     # MCPB manifest
├── README.md                         # End-user documentation
└── icon-512.png                      # Server icon (512x512)
```

## Release Workflow

1. **Create GitHub Release:**
   - Tag format: `v1.x.x`
   - Upload `PptMcp.McpServer-win-x64.zip` as release asset

2. **Update manifest.json download URL:**
   - Verify the `install.win32.download` URL points to the release asset
   - URL format: `https://github.com/trsdn/mcp-server-ppt/releases/latest/download/PptMcp.McpServer-win-x64.zip`

3. **Submit to Claude Directory:**
   - Follow Anthropic's submission process
   - Include the manifest.json content

## Manifest Schema

The manifest follows MCPB version 0.3 specification:

```json
{
  "manifestVersion": "0.3",
  "server": {
    "id": "ppt-mcp-server",
    "name": "PowerPoint MCP Server",
    "type": "binary",
    "platforms": ["win32"]
  },
  "install": {
    "win32": {
      "download": "https://github.com/.../PptMcp.McpServer-win-x64.zip",
      "command": "PptMcp.McpServer.exe"
    }
  }
}
```

## Tool Annotations

All 22 MCP tools include the `Destructive = true` annotation since they can modify PowerPoint files:

```csharp
[McpServerTool(Name = "range", Title = "PowerPoint Range Operations", Destructive = true)]
```

## Technical Notes

### Why Self-Contained?

- Users don't need .NET SDK installed
- Avoids version conflicts
- Single executable deployment (~15 MB compressed)

### Why No Trimming?

PowerPoint COM interop uses `Type.GetTypeFromProgID()` which requires reflection. Trimming would break COM activation with IL2072 errors.

### Why Windows x64 Only?

- COM interop requires Windows
- x64 is the most common architecture
- ARM64 Windows can run x64 binaries via emulation

## Verification

After building, verify the package:

```powershell
# List zip contents
Expand-Archive ./artifacts/PptMcp.McpServer-win-x64.zip -DestinationPath ./test-extract
dir ./test-extract
Remove-Item -Recurse ./test-extract
```

## Submission Guidelines Reference

See the [MCPB Submission Guide](https://support.claude.com/en/articles/12922832-local-mcp-server-submission-guide) for:
- Tool annotation requirements (readOnlyHint, destructiveHint)
- README requirements (minimum 3 examples with expected behavior)
- Privacy policy requirements
- manifest_version requirements (≥ 0.3)

## Links

- [MCPB Specification](https://modelcontextprotocol.io/docs/registry)
- [Claude Desktop Documentation](https://docs.anthropic.com/claude/docs/claude-for-desktop)
- [Submission Guidelines](https://support.claude.com/en/articles/12922832-local-mcp-server-submission-guide)
