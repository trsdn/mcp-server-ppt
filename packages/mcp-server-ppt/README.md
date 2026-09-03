# mcp-server-ppt

MCP server for Microsoft PowerPoint automation through PowerPoint's own COM API.

```powershell
npx -y mcp-server-ppt
```

No .NET SDK or runtime required — the wrapper downloads a self-contained build.

## Requirements

- **Windows x64.** `npm` refuses to install this package anywhere else, because
  COM interop does not exist on other platforms.
- **Microsoft PowerPoint 2016 or later**, installed locally. This server drives the
  real application; it is not a file-format library.

## Client configuration

VS Code (`.vscode/mcp.json`) or Visual Studio (`.mcp.json`):

```json
{
  "servers": {
    "ppt-mcp": {
      "command": "npx",
      "args": ["-y", "mcp-server-ppt"]
    }
  }
}
```

Claude Desktop, Cursor, Cline, Windsurf:

```json
{
  "mcpServers": {
    "ppt-mcp": {
      "command": "npx",
      "args": ["-y", "mcp-server-ppt"]
    }
  }
}
```

## How the install works

On first run the wrapper downloads `PptMcp-MCP-Server-<version>-win-x64.zip` from the
matching GitHub release and extracts it to
`%LOCALAPPDATA%\mcp-server-ppt\runtime-<version>`. Subsequent runs reuse it. The
download is roughly 64 MB compressed, 145 MB once extracted, and happens once per
version.

Downloading lazily rather than in a `postinstall` hook is deliberate: npm 12 blocks
dependency install scripts by default, and `npx` installs this package as a dependency
of a temporary root, so a `postinstall` hook is not something that can be relied on.
Resolving inside `bin` always runs.

To download ahead of time instead of on the first MCP request:

```powershell
npx -y mcp-server-ppt --install
```

If extraction or download fails, the partially populated directory is removed rather
than left behind — otherwise the next run would find no executable, re-download, and
`Expand-Archive -Force` would merge the two trees.

### Environment variables

| Variable | Purpose |
| --- | --- |
| `MCP_SERVER_PPT_HOME` | Use an existing extracted build at this path and skip the download entirely. |
| `MCP_SERVER_PPT_CACHE` | Base directory for the cached runtime. Defaults to `%LOCALAPPDATA%\mcp-server-ppt`. |
| `MCP_SERVER_PPT_ASSET_URL` | Download from somewhere other than the GitHub release, for mirrors and testing. |

## Other ways to install

- **.NET tool:** `dotnet tool install --global PptMcp.McpServer`
- **VS Code extension**, **Claude Desktop MCPB bundle**, and direct release downloads:
  see the [main README](https://github.com/trsdn/mcp-server-ppt).

## Licence

MIT. Source: https://github.com/trsdn/mcp-server-ppt
