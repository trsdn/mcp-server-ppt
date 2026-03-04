# Claude Desktop Configuration

PowerPoint MCP Server works with Claude Desktop on Windows, but requires specific configuration for the Windows container environment.

## Configuration Location

Claude Desktop config file:
```
%APPDATA%\Claude\claude_desktop_config.json
```

## Basic Configuration

```json
{
  "mcpServers": {
    "ppt-mcp": {
      "command": "ppt-mcp-server.exe",
      "args": []
    }
  }
}
```

Or using the .NET tool:
```json
{
  "mcpServers": {
    "ppt-mcp": {
      "command": "dotnet",
      "args": ["ppt-mcp-server"]
    }
  }
}
```

## Windows Container Considerations

Claude Desktop runs in a Windows container with specific constraints:

### File System Access

The container has limited file system access. PowerPoint files should be in accessible locations:

- **User Documents**: `C:\Users\<username>\Documents\`
- **User Desktop**: `C:\Users\<username>\Desktop\`
- **Temp directory**: `%TEMP%` or `C:\Users\<username>\AppData\Local\Temp\`

**Recommendation**: Work with files in your Documents folder.

### PowerPoint Instance

- PowerPoint MCP Server manages its own PowerPoint instance via COM automation
- The PowerPoint window may be visible or hidden depending on operation
- Long-running operations show PowerPoint's progress indicators

### Session Persistence

Sessions are tied to the Claude Desktop session:
- Closing Claude Desktop terminates active PowerPoint sessions
- Unsaved changes may be lost
- Use explicit `file(action: 'close', save: true)` to persist work

## Recommended Workflow

```
1. Create or open file in accessible location:
   file(action: 'create', filePath: 'C:\\Users\\Me\\Documents\\report.pptx')

2. Perform operations with returned sessionId

3. Explicitly save and close when done:
   file(action: 'close', sessionId: '...', save: true)
```

## Troubleshooting

### "PowerPoint not found" Error
- Ensure Microsoft PowerPoint is installed on the Windows system
- PowerPoint 2016, 2019, 2021, or Microsoft 365 required

### "Access denied" Error
- Check file path is in accessible directory
- Ensure file is not open in another PowerPoint instance
- Try using Documents folder instead of other locations

### "COM timeout" Error
- PowerPoint may be showing a dialog - check for visible PowerPoint window
- Operation may be long-running - wait for completion
- Restart Claude Desktop if PowerPoint becomes unresponsive

### VBA Operations Fail
VBA requires explicit trust setting in PowerPoint:
1. Open PowerPoint Options → Trust Center → Trust Center Settings
2. Enable "Trust access to the VBA project object model"
3. Restart PowerPoint MCP Server

## MCPB Bundle Alternative

For simplified installation, use the MCPB bundle which auto-configures Claude Desktop:

1. Download `ppt-mcp-bundle.mcpb` from releases
2. Double-click to install
3. Restart Claude Desktop

See the main repository for MCPB installation instructions.
