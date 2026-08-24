# Privacy

PptMcp is a Windows desktop tool. It automates a locally installed copy of Microsoft
PowerPoint through COM and runs entirely on your machine.

This document describes what the software does with your data. It is referenced by
`mcpb/manifest.json`, which desktop MCP clients surface to users before installation.

## Your presentations stay on your machine

PptMcp opens, edits and saves the `.pptx` / `.pptm` files you point it at, using the
PowerPoint installation on the same computer. Slide content, speaker notes, embedded
data and file paths are not uploaded anywhere by this software.

There is no telemetry, no analytics and no crash reporting.

Note that PptMcp is normally driven by an AI assistant (an MCP client such as Claude
Desktop or VS Code, or a coding agent using the CLI). Whatever that assistant chooses
to send to its own model provider is governed by that assistant's privacy policy, not
this one. If you ask an assistant to read a slide and it repeats the text back to you,
that text passed through the assistant's provider.

## The one network request

PptMcp makes a single outbound request, to check whether a newer release exists:

```
GET https://api.nuget.org/v3-flatcontainer/pptmcp.mcpserver/index.json   (MCP Server)
GET https://api.nuget.org/v3-flatcontainer/pptmcp.cli/index.json         (CLI)
```

- It runs at startup, and for the CLI also from the system tray.
- It sends no data about you, your files or your usage - it is a plain GET for the
  list of published versions. As with any HTTP request, nuget.org will see your IP
  address and the request metadata. That is covered by the
  [NuGet privacy statement](https://www.nuget.org/policies/Privacy).
- If the request fails, it is ignored and the tool continues normally.
- The request times out after 5 seconds.

There is currently **no setting to disable this check**. Blocking `api.nuget.org` at
the network level is safe: the check fails silently and nothing else depends on it.
See [VERSION-CHECKING.md](VERSION-CHECKING.md) for the implementation.

## Macros and VBA

Some operations write and execute VBA inside the presentation you are working on.
That code runs locally under your Windows user account with the permissions PowerPoint
has - the same as if you had written the macro yourself. Only run PptMcp against
presentations you trust.

## Questions

Open an issue at
[github.com/trsdn/mcp-server-ppt/issues](https://github.com/trsdn/mcp-server-ppt/issues).
