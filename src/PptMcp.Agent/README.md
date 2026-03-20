# PptMcp.Agent

Official source-side Copilot SDK orchestrator for multi-phase PowerPoint deck generation through `mcp-server-ppt`.

## Purpose

`src\PptMcp.Agent` is the repository's first-class deck-building controller.

It exists for workflows that are larger than a single prompt/response exchange:

- plan a deck from one natural-language task
- execute the plan through normal sequential MCP tool calls
- verify the produced presentation
- repair incomplete output when artifact and quality validation fail

This keeps the product boundary clean:

- `src\PptMcp.McpServer` stays focused on primitive PowerPoint capabilities
- `skills\shared\*.md` stays focused on LLM guidance
- `src\PptMcp.Agent` owns orchestration, retries, artifact validation, and run summaries

## Architectural Boundary

`PptMcp.Agent` is **not** a third server surface and it does **not** move orchestration into the MCP server.

It deliberately avoids:

- MCP batch dependencies
- MCP subagent dependencies
- server-side planner / worker / verifier state machines

Instead, one client process runs these logical phases:

1. **Plan** — generate structured slide intents without touching PowerPoint
2. **Execute** — build the deck through standard MCP calls
3. **Verify** — reopen and inspect the generated deck
4. **Repair** — re-enter the deck if structural or business-quality validation finds gaps

## What the Agent Writes

For an output file like `quarterly-review.pptx`, the agent also writes:

- `quarterly-review.plan.json` — extracted structured slide plan
- `quarterly-review-artifacts\` — verification exports and runtime traces
- `quarterly-review-artifacts\run-summary.json` — high-level execution summary

## Run from Source

```powershell
dotnet build src\PptMcp.McpServer\PptMcp.McpServer.csproj -c Release

Set-Location src\PptMcp.Agent
npm install
npm run check
npm test

node .\src\cli.mjs run `
  --task "Build a 5-slide executive deck on Q4 revenue performance and next actions." `
  --output "C:\Users\you\Documents\q4-revenue-deck.pptx"
```

Optional flags:

```powershell
--model gpt-5.4
--plan-file C:\path\to\precomputed.plan.json
--show
--overwrite
--skip-verify
--mcp-server "C:\path\to\PptMcp.McpServer.exe"
--plan-timeout-ms 120000
--execute-timeout-ms 900000
--verify-timeout-ms 300000
```

## Default MCP Server Resolution

By default the client looks for:

- `src\PptMcp.McpServer\bin\Release\net9.0-windows\PptMcp.McpServer.exe`

You can override that with:

- `--mcp-server`
- `PPT_MCP_AGENT_MCP_SERVER`
- `PPT_MCP_SERVER_COMMAND`
- `ppt_mcp_SERVER_COMMAND`

## Reusing a Precomputed Plan

If planning is already done or you want to debug execution in isolation, you can skip the planning phase:

```powershell
node .\src\cli.mjs run `
  --task "Execute this dashboard plan." `
  --plan-file "C:\path\to\dashboard.plan.json" `
  --output "C:\Users\you\Documents\dashboard.pptx"
```

Plan-file runs still validate the saved deck against required literal slide text from the plan. They now also run deterministic business-quality checks that reject novelty shapes like `sun` on dashboard-style slides and flag overly vivid palettes for business content. If execution times out after producing a partial or low-quality artifact, the agent reopens the deck, repairs it, and validates again before succeeding.

## Reliability Behaviors

The official source component keeps the hardened behaviors from the verified prototype:

- JSON-first plan extraction with markdown fallback
- timeout fallback when the expected artifact already exists
- save-before-close cleanup for open agent presentations
- ZIP-based PPTX slide-count validation
- literal text validation against the plan
- deterministic business-quality validation for business slides
- targeted repair pass when the generated deck is incomplete or visually unacceptable

## Validation Scope

Local validation for this component should always cover:

```powershell
Set-Location src\PptMcp.Agent
npm run check
npm test
```

For end-to-end smoke validation, build the MCP server and run a small `run --task ...` scenario against a real PowerPoint-enabled Windows desktop.

## Related Docs

- [Agent Client Architecture](../../docs/AGENT-CLIENT.md)
- [Installation Guide](../../docs/INSTALLATION.md)
- [Eval Framework](../../eval/README.md)
- [Archetype Pipeline](../../docs/ARCHETYPE-PIPELINE.md)
