# Agent Client Architecture

`src\PptMcp.Agent` is the official source-side Copilot SDK controller for multi-phase PowerPoint deck generation in this repository.

It exists to handle orchestration that should **not** live inside the MCP server itself.

## Why This Component Exists

Some deck-generation tasks need more than one prompt:

- create a structured plan first
- build the deck slide-by-slide
- re-open and verify the result
- repair incomplete or visually unacceptable output if the produced artifact is wrong

That logic is intentionally client-side.

The MCP server remains responsible for primitive PowerPoint capabilities, while the agent owns sequencing, retries, artifact checks, and run summaries.

## Responsibility Split

| Layer | Responsibility |
|---|---|
| `src\PptMcp.Core` + `src\PptMcp.Service` + `src\PptMcp.McpServer` | Primitive PowerPoint operations, sessions, and MCP transport |
| `skills\shared\*.md` | Planning, design, and review guidance shared across hosts |
| `src\PptMcp.Agent` | Plan → execute → verify → repair orchestration plus deterministic artifact/quality validation |
| `eval\` | Experimental measurement, scoring, sweeps, and skill-tuning loops |

## Design Constraints

The component deliberately avoids pushing orchestration into the server:

- no MCP batch dependency
- no MCP subagent dependency
- no server-side planner / worker state machine
- no hidden session coordinator inside the MCP service

Instead, one local process controls the entire workflow and talks to the MCP server with normal sequential tool calls.

## Runtime Flow

The orchestrator currently works in four logical phases.

### 1. Plan

- Reads archetype guidance from `src\PptMcp.Core\Data\archetypes\registry.md`
- Reads generation guidance from `skills\shared\generation-pipeline.md`
- Asks the model for a JSON slide plan
- Falls back to fenced JSON, outermost JSON, or markdown slide blocks if the reply is not perfectly structured

Output:

- `*.plan.json`

### 2. Execute

- Creates a new presentation at the requested output path
- Builds slides in plan order
- Uses normal MCP tool calls only
- Prefers placeholders when the layout already exposes them
- Requires `file(action='close', save=true)` before finishing

### 3. Verify

- Re-opens the generated deck
- Inspects structure with standard MCP read/list operations
- Can export slide images into the artifact directory for review
- Applies targeted fixes for structural and visual business-quality problems

### 4. Repair

- Runs if artifact validation detects an incomplete or low-quality result
- Re-opens or recreates the deck
- Repairs the structure against the fixed plan
- Verifies the final slide count before saving and closing

## Reliability Behaviors

The official client keeps the hardened behaviors from the verified prototype:

- **Plan parsing fallback** — handles JSON objects, nested `plan.slides`, arrays, fenced JSON, and markdown slide blocks
- **Timeout fallback** — if the SDK times out waiting for `session.idle` but the expected artifact already exists, the phase can still succeed
- **Save-before-close cleanup** — targeted cleanup saves open presentations before closing them
- **Artifact validation** — PPTX output is reopened as a zip and checked for slide XML entries plus required literal plan text
- **Quality validation** — business-oriented slides are rejected when they contain novelty preset shapes or overly vivid palettes
- **Repair loop** — incomplete or visually weak output triggers a repair phase instead of silently accepting a broken deck

## Files in This Component

| Path | Purpose |
|---|---|
| `src\PptMcp.Agent\src\cli.mjs` | CLI entry point |
| `src\PptMcp.Agent\src\orchestrator.mjs` | Phase sequencing and artifact validation |
| `src\PptMcp.Agent\src\runtime.mjs` | Copilot SDK session/runtime wrapper |
| `src\PptMcp.Agent\src\planner.mjs` | Plan extraction and normalization |
| `src\PptMcp.Agent\patch-deps.cjs` | Node 24 compatibility patch for `@github/copilot-sdk` dependencies |
| `src\PptMcp.Agent\tests\planner.test.mjs` | Fast local regression coverage for plan parsing |

## Build and Test

```powershell
dotnet build src\PptMcp.McpServer\PptMcp.McpServer.csproj -c Release

Set-Location src\PptMcp.Agent
npm install
npm run check
npm test
```

For an end-to-end smoke run on a PowerPoint-enabled Windows desktop:

```powershell
node .\src\cli.mjs run `
  --task "Build a one-slide title presentation with the title 'Agent Smoke Test' and subtitle 'Prototype validation'." `
  --output "C:\Users\you\AppData\Local\Temp\ppt-agent-smoke.pptx" `
  --overwrite
```

## Output Artifacts

For `deck.pptx`, the client also writes:

- `deck.plan.json`
- `deck-artifacts\run-summary.json`
- optional review/export artifacts inside `deck-artifacts\`

These artifacts make the run inspectable and reproducible without pushing orchestration into the server.

## Relationship to the Eval Framework

The agent and eval harnesses use related building blocks, but they are not the same thing:

- `src\PptMcp.Agent` is for one production-style build workflow
- `eval\` is for repeated experiments, judgments, and score histories

See also:

- [Eval Framework](../eval/README.md)
- [Archetype Pipeline](ARCHETYPE-PIPELINE.md)
