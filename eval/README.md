# Slide Eval Framework

Config-driven evaluation harnesses for testing and improving PptMcp slide-building behavior with real PowerPoint outputs.

This directory is a development-only workflow. It is not part of the shipped MCP package.

## Related Documentation

- [Official Agent Client](../src/PptMcp.Agent/README.md)
- [Agent Client Architecture](../docs/AGENT-CLIENT.md)
- [Archetype Pipeline](../docs/ARCHETYPE-PIPELINE.md)

## Eval vs Official Agent Client

The repo now contains both an official source-side agent client and the eval harnesses in this folder. They solve different problems:

| Area | Primary goal | Typical output |
|---|---|---|
| `src\PptMcp.Agent` | Build a deck from one user task through plan → execute → verify → repair | A presentation plus plan and run artifacts |
| `eval\` | Measure, compare, and improve builder behavior across repeated scenarios | Score histories, PNGs, manifests, and run reports |
| `skills\shared\*.md` | Shared guidance consumed by both | Reusable planning and design rules |

Use `src\PptMcp.Agent` when you want one production-style build workflow.

Use `eval\` when you want repeatable experiments, score tracking, archetype sweeps, or skill-tuning loops.

## Reference-Slide and Archetype Pipeline Overview

The eval framework is also where reference-slide curation and learned-archetype material starts:

1. Extract or import individual slides into `eval\input\individual-slides`
2. Triage them with the direct LLM workflow into `eval\output\slide-triage\good` and `eval\output\slide-triage\reject`
3. Use batch-list and batch-result artifacts to normalize and classify the accepted slides
4. Regenerate the sanitized learned-reference catalog consumed by the runtime design surface

The full privacy and rebuild rules live in [Archetype Pipeline](../docs/ARCHETYPE-PIPELINE.md).

## Current Status

The eval area currently supports two related but different workflows:

1. **`run-archetype-eval.mjs`** — the current primary harness for archetype-by-archetype evaluation, optional skill tuning, MCP transport experiments, session reuse, and isolated process runs.
2. **`run-eval.mjs`** — an older prompt-sweep harness that is still useful for broad CLI-based checks and builder/judge comparisons, but it is not the main direction for skill-tuning architecture.

Use `run-archetype-eval.mjs` for most new work.

## Implemented Architecture Layers

The current archetype harness is organized around a few concrete layers under `eval/lib/`:

- `mode.mjs` — resolves the practical `baseline` vs `tuning` split and scopes run names, `output/`, and `results/` directories.
- `runtime/` — creates Copilot SDK runtimes, handles fresh vs reused sessions, and supports isolated child-process execution.
- `protocol/` — defines strict envelopes such as `evaluation-request/v1`, `builder-summary/v1`, and `judge-response/v1`, plus typed failure categories.
- `orchestrator/` — runs the loop phases in order: cleanup, build, artifact verification, judge, optional improve, and recovery/finalization.
- `persistence/` — writes per-run ledgers, manifests, transaction logs, artifact manifests, and skill snapshots.
- `reporting/` — builds the operator-facing `run-report.json` summary from persisted loop records.

This is the architecture that exists in code today. It is more structured than the older prompt-sweep workflow, but it is still not the final ACP-based rewrite.

## What Is Implemented Today

### Primary loop in `run-archetype-eval.mjs`

Per loop, the harness can:

- build one slide from a config prompt
- export a PNG artifact
- judge the exported image against a structural rubric
- optionally run an improver step against `skills/shared/*.md`
- save structured per-loop results to `eval/results/...`

### Supported execution controls

The current harness already supports the following concepts:

| Capability | Current support | Notes |
|---|---|---|
| `reasoningEffort` | Yes | Passed through to Copilot SDK sessions when configured. |
| MCP transport | Yes | Supported in `run-archetype-eval.mjs` for builder sessions via `builder.transport: "mcp"`. Default builder path remains CLI-oriented when omitted. |
| isolated process mode | Yes | `isolatedProcess: true` runs a fresh Copilot child process for that agent step. |
| session reuse | Yes | `reuseSessionContext: true` keeps one session alive across loops, unless isolated mode is enabled. |
| explicit JSON reviewer output | Yes | Judge instructions require one JSON object only; the archetype harness normalizes that JSON into structured result fields. |
| explicit builder follow-up summary | Yes | After a successful build, the archetype harness asks the builder for a JSON summary and stores it alongside the raw response. |
| bounded retries | Yes | Build and judge steps retry once for typed transient failures such as runtime drops, malformed review JSON, and unstable artifacts. |
| transport/instruction validation | Yes | Builder transport is validated against the configured instructions file before execution. |
| artifact readiness validation | Yes | PNG/PPTX artifacts are checked for size, signature, and short stability before the judge runs. |
| runtime selector | Partial | `copilot-sdk` is implemented; `acp` is reserved and currently throws an invalid-configuration error. |
| persistence ledger | Yes | Each run writes a manifest, transaction ledger, per-loop records, artifact manifests, and skill snapshots. |
| run report generation | Yes | Each run writes `run-report.json` with operator-oriented summary metrics and alerts. |

### Important limitation

The current system is **not yet the full rewritten eval architecture** described in the session plan. In particular:

- there is now a **practical mode split** between baseline eval and skill-tuning eval, but not yet a full end-to-end architecture rewrite
- the only implemented runtime is **Copilot SDK**; ACP wiring is not available yet
- there is **not yet a generalized transport abstraction layer**
- reviewer JSON is strongly encouraged and normalized, but the harness still contains fallback parsing paths for degraded cases
- strict response contracts now live under `eval/lib/protocol/`; fallback parsing still exists, but it is recorded explicitly as a typed contract failure instead of being silent
- mode-specific builder behavior is now implemented, but the broader architecture is still evolving

So treat the current system as a practical, stability-focused harness with several forward-looking features already in place, not as the final architecture endpoint.

## Stability-Oriented Usage Guidance

### When to use fresh sessions, reuse, or isolated mode

| Mode | Best for | Why | Trade-off |
|---|---|---|---|
| **Fresh session per loop** (default) | Normal eval runs | Good balance of simplicity and independence | Less carry-over context |
| **`reuseSessionContext: true`** | Skill-tuning loops where prior reviewer feedback should influence the next build | Lets the builder keep prior conversation state across loops | More hidden state, weaker apples-to-apples comparisons |
| **`isolatedProcess: true`** | Stability testing, strict A/B runs, cleaner reproducibility, crash containment | Hard isolation between loops or judged slides | Higher startup cost, no conversational carry-over |

### Reuse vs isolation

Do **not** think of reuse and isolation as interchangeable:

- Use **reuse** when you want deliberate continuity inside one archetype cycle.
- Use **isolation** when you want each run to stand on its own with minimal hidden state.
- If both are configured, isolation wins in the current code path.

### Recommended defaults by objective

| Objective | Recommended pattern |
|---|---|
| Baseline model comparison | Freeze skills manually, keep configs separate, prefer fresh or isolated runs, avoid improver edits |
| Skill tuning | Use `run-archetype-eval.mjs`, enable improver if needed, optionally reuse builder session context |
| Transport comparison | Keep prompts and skills fixed, vary `builder.transport`, and prefer fresh or isolated execution |
| Reliability debugging | Prefer `isolatedProcess: true` to narrow session contamination and cleanup issues |

## Baseline Eval vs Skill-Tuning Direction

This distinction matters:

- **Baseline model eval** means measuring model/runtime behavior against a fixed skill set.
- **Skill tuning** means changing `skills/shared/*.md` based on reviewer feedback and then measuring the next loop.

### Explicit eval modes

`run-archetype-eval.mjs` now supports an explicit config-level `mode`:

- `baseline` — compare models/runtimes against a fixed skill snapshot
- `tuning` — allow iterative skill changes and score tracking inside that tuning run

If `mode` is omitted, the harness keeps older behavior:

- configs **with** an `improver` section default to `tuning`
- configs **without** an `improver` section default to `baseline`

Important behavior:

- baseline mode always writes under `output/.../baseline` and `results/.../baseline`
- tuning mode always writes under `output/.../tuning` and `results/.../tuning`
- baseline mode ignores any configured `improver`
- persisted manifests, loop records, JSON summaries, and run names now include the mode

`run-eval.mjs` is baseline-oriented by default and now also accepts `--mode baseline|tuning` so manual experiments can keep artifacts and score histories separated.

## High-Level Flow

### Archetype harness (`run-archetype-eval.mjs`)

```text
Config
  -> builder builds PPTX/PNG
  -> judge inspects PNG and returns JSON
  -> harness normalizes/stores structured results
  -> optional improver edits skills
  -> next loop
```

Key runtime behavior today:

- builder is fresh per loop by default
- builder can reuse one session across loops
- builder can run in isolated child-process mode
- judge uses one persistent session across the archetype run by default
- judge can also be isolated
- improver is optional and runs only when configured and gaps exist
- cleanup, build, artifact, and judge failures are recorded with typed categories plus retry metadata

### Prompt-sweep harness (`run-eval.mjs`)

This harness remains useful for:

- broad prompt sweeps
- older CLI-oriented eval runs
- fast builder/judge comparison checks

It also supports:

- builder/judge `reasoningEffort`
- builder/judge session reuse
- builder/judge isolated child-process mode

But it is still the older workflow and should not be treated as the long-term skill-tuning architecture.

## Setup

### Prerequisites

- Node.js 18+
- GitHub Copilot CLI installed and authenticated
- PptMcp CLI built: `dotnet build src/PptMcp.CLI -c Release`
- For MCP builder transport: build the MCP server too: `dotnet build src/PptMcp.McpServer -c Release`
- Windows with PowerPoint installed

### Install

```bash
cd eval
npm install
npm run postinstall
```

### Start the CLI service when using CLI-backed builds

```powershell
pptcli service start
```

The archetype harness calls service startup defensively, but having the CLI service available up front is still the safest workflow for CLI-backed runs and fallback export paths.

## Private Eval Asset Root

The eval code stays in this repo, but the large/private assets can live in a separate private repo clone.

Set `PPTMCP_EVAL_ASSET_REPO_ROOT` to the **root of that private repo**. When set, the harnesses read and write eval assets under:

- `<asset-repo>\eval\input`
- `<asset-repo>\eval\output`
- `<asset-repo>\eval\results`
- `<asset-repo>\eval\data\archetype-references`

When the variable is not set, the current repo root remains the default asset root.

`PPTMCP_REFERENCE_DATA_ROOT` still works as a direct override for the reference catalog itself, but the preferred repo-local path is now `eval\data\archetype-references` instead of `local-data\archetype-references`.

To move assets between this repo and a private repo clone, use:

```powershell
.\scripts\Sync-EvalAssets.ps1 -Direction Import -PrivateRepoRoot C:\path\to\private-eval-repo
.\scripts\Sync-EvalAssets.ps1 -Direction Export -PrivateRepoRoot C:\path\to\private-eval-repo
```

The sync script copies `eval\input`, `eval\output`, `eval\results`, and `eval\data`, while excluding `eval\input\master.pdf` and `eval\input\extract*`.

## Running Evaluations

### Preferred: archetype harness

```bash
node run-archetype-eval.mjs configs\kpi-dashboard.json
node run-archetype-eval.mjs configs\title-slide.json
node run-archetype-eval.mjs configs\all
```

### Direct reference-slide triage

Use this when you want the LLM to inspect already extracted slide PNGs directly, without MCP or PowerPoint build steps, and sort them into strong reference examples vs rejects.

```bash
node triage-slides.mjs
node triage-slides.mjs --limit 25 --dry-run
node triage-slides.mjs --model claude-sonnet-4.5 --min-score 16
node triage-slides.mjs --shard-count 8 --shard-index 0
```

Default behavior:

- reads PNGs from `<asset-root>\eval\input\individual-slides`
- scores each slide directly via Copilot SDK using the judge prompt
- applies a strict accept threshold focused on structure and visual quality
- allows non-action-title cover/admin slides to pass if they are otherwise strong
- supports deterministic sharding via `--shard-count` / `--shard-index` for parallel direct-LLM workers
- moves accepted slides to `<asset-root>\eval\output\slide-triage\good`
- moves rejected slides to `<asset-root>\eval\output\slide-triage\reject`
- writes per-slide JSONL results and a rolling summary under `<asset-root>\eval\output\slide-triage`

### Prompt sweep harness

```bash
node run-eval.mjs
node run-eval.mjs --mode baseline
node run-eval.mjs --mode tuning
node run-eval.mjs --start 0 --count 5
node run-eval.mjs --category dashboard
```

## Config Schema for `run-archetype-eval.mjs`

Each JSON file in `eval/configs/` is the source of truth for one archetype run.

```jsonc
{
  "name": "title-slide-ab-opus46-high-reuse",
  "mode": "baseline",
  "description": "What this run is measuring",
  "goal": "Operator-facing success target",
  "archetype": "title-slide",
  "loops": 4,

  "builder": {
    "model": "claude-opus-4.6",
    "reasoningEffort": "high",            // optional
    "transport": "mcp",                   // optional; omit for default CLI-style path
    "reuseSessionContext": true,          // optional
    "isolatedProcess": false,             // optional
    "executionMode": "reuse-session",     // optional alias in current code
    "skillFiles": [
      "slide-design-principles.md",
      "slide-design-review.md"
    ],
    "instructionsFile": "agents/builder-instructions-mcp.md",
    "timeoutMs": 600000                   // optional; defaults depend on model/effort
  },

  "judge": {
    "model": "gpt-5.4",
    "reasoningEffort": "medium",          // optional
    "isolatedProcess": false,             // optional
    "executionMode": "isolated-process",  // optional alias in current code
    "instructionsFile": "agents/judge-instructions.md",
    "timeoutMs": 120000
  },

  "improver": {
    "model": "claude-sonnet-4",
    "reasoningEffort": "medium",
    "isolatedProcess": true,
    "instructionsFile": "agents/improver-instructions.md",
    "timeoutMs": 180000
  }, // only used in tuning mode; ignored in baseline

  "prompts": [
    { "id": "title-board", "text": "Build a title slide for..." }
  ],

  "outputDir": "output/title-slide-ab-opus46-high-reuse",
  "resultsDir": "results/title-slide-ab-opus46-high-reuse"
}
```

## Config Notes

| Setting | What it means in the current code |
|---|---|
| `builder.reasoningEffort` | Passed directly into Copilot session config. |
| `builder.transport: "mcp"` | Builder session gets the PowerPoint MCP server attached. |
| `runtime` | Optional runtime selector for the shared layer under `eval/lib/runtime/`. Omit it for the current Copilot SDK path; ACP is reserved for future work and not implemented yet. |
| `reuseSessionContext` | Keeps one session/runtime alive across loops unless isolation is enabled. |
| `isolatedProcess` | Uses `copilot-isolated-worker.mjs` to run the step in a separate child process. |
| `mode` | `baseline` keeps skills fixed and disables the improver; `tuning` allows improver-driven skill edits. Omit it to infer the older behavior from whether `improver` is configured. |
| `executionMode` | Alternate way to express `reuse-session` or `isolated-process` in the archetype harness. Most useful for builder and improver settings; the judge is already persistent unless isolated. |
| `skillFiles` | Files the builder is told to read from `skills/shared/`. |
| `instructionsFile` | Agent-specific instruction file under `eval/agents/`. |
| `timeoutMs` | Explicit override; otherwise builder defaults scale with model and reasoning effort. |

### Runtime and execution-mode rules

- If `runtime` is omitted, the harness uses `copilot-sdk`.
- `runtime: "acp"` is **not implemented** and currently fails fast.
- `isolatedProcess: true` wins over any reuse setting.
- The archetype harness keeps the **judge** in one reused session by default unless `judge.isolatedProcess` is enabled.
- The archetype harness keeps the **builder** fresh by default unless `builder.reuseSessionContext` or `builder.executionMode: "reuse-session"` is enabled.

### Builder timeout defaults

If `builder.timeoutMs` is omitted in `run-archetype-eval.mjs`, the harness currently uses:

- 5 minutes by default
- 7 minutes for Opus
- 8 minutes for `high`
- 10 minutes for `xhigh`, or `high` with Opus

Use explicit `timeoutMs` when you want stable comparisons across models.

## Judge Contract

`eval/agents/judge-instructions.md` is the reviewer contract.

The current expectation is:

- inspect the exported PNG directly
- score the 7 structural dimensions
- return **one JSON object only**
- include `dimensionScores`, `totalScore`, `maxScore`, and `gaps`

The archetype harness stores:

- normalized score fields
- per-dimension scores
- judge summary
- expected vs used archetype
- raw judge output for auditing

The preferred reviewer response is now the strict `judge-response/v1` envelope. If the model falls back to legacy JSON or prose-like output, the harness records a typed validation result (`schema_error` or `review_invalid`) and only uses degraded parsing when it can do so explicitly.

## Builder Contract

The builder instructions differ by transport:

- `agents/builder-instructions.md` — CLI-oriented builder flow
- `agents/builder-instructions-mcp.md` — MCP tool flow

For archetype runs, the harness also asks the builder for a short JSON follow-up summary after a successful build. That summary is normalized and stored in results as:

- archetype
- palette
- shape count
- preserved facts
- rationale

That follow-up summary now uses the `builder-summary/v1` envelope. If the builder returns legacy JSON or unstructured text, the harness keeps running but stores the fallback validation status in the result JSON.

## Results and Persistence Layout

Both harnesses now write **mode-scoped artifacts** and **mode-scoped ledgers**.

### Archetype harness outputs

- slide artifacts: `outputDir/<mode>/loop{N}-{promptId}.png` and `.pptx`
- top-level summary JSON: `resultsDir/<mode>/<runName>-<timestamp>.json`
- persistence ledger root: `resultsDir/<mode>/ledger/<runId>/`
- generated operator report: `resultsDir/<mode>/ledger/<runId>/run-report.json`

### Prompt-sweep harness outputs

- slide artifacts: `<asset-root>/eval/output/<mode>/auto-<promptId>.png` and `.pptx`
- top-level summary JSON: `<asset-root>/eval/results/<mode>/<runName>-<timestamp>.json`
- persistence ledger root: `<asset-root>/eval/results/<mode>/ledger/<runId>/`
- generated operator report: `<asset-root>/eval/results/<mode>/ledger/<runId>/run-report.json`

### What lives inside the ledger

```text
results/<mode>/ledger/<runId>/
├── manifest.json
├── transactions.jsonl
├── run-report.json
└── loops/
    ├── loop-0001-<promptId>.json
    ├── loop-0002-<promptId>.json
    └── ...
```

Important persistence details:

- `manifest.json` is the run-level record (`eval-run-manifest/v1`) with counters, metadata, and final artifacts.
- `transactions.jsonl` is the append-only ledger of loop outcomes.
- each `loops/loop-*.json` file is a full structured loop record (`eval-loop-record/v1`)
- artifact fingerprints and skill snapshots are stored inside those loop records via dedicated manifests
- `run-report.json` is rebuilt from persistence and is the best operator-facing entry point after a run finishes

Typical loop-level fields include:

- `build_ok`
- `judge_ok`
- `score`
- `max_score`
- `gaps`
- `judge_summary`
- `judge_archetype_used`
- `judge_archetype_expected`
- `judge_dimension_scores`
- `builder_summary`
- `builder_raw`
- `judge_raw`
- `png_path`
- optional `improvement`

This makes the archetype harness the better choice when you need auditable loop history rather than just top-line averages.

## Operator Checklist: Where to Look First

After a run finishes, use this order:

1. `run-report.json` — operator summary, repeated gaps, retry counts, failure categories, and score regression alerts
2. `manifest.json` — run metadata, counters, and final artifact list
3. `loops/loop-*.json` — per-loop raw builder/judge output, validation status, artifact fingerprints, and skill snapshots
4. top-level results JSON — convenient bundle containing the summarized report plus the flattened `results` array

If you are debugging instability or bad scores, inspect `builder.validation`, `judge.validation`, `diagnostics.recovery`, and the artifact manifest inside the relevant loop record.

## Agent Instruction Files

### `agents/builder-instructions.md`

CLI-oriented builder guidance:

- read design skills
- build one slide with `pptcli`
- export PNG
- save and close
- stay available for a `builder-summary/v1` follow-up request

### `agents/builder-instructions-mcp.md`

MCP-oriented builder guidance:

- use PowerPoint MCP tools directly
- prefer file -> slide -> shape/text -> export -> close
- stay in the same conversation so the harness can request a structured JSON summary

### `agents/judge-instructions.md`

Reviewer guidance:

- structure and archetype correctness only
- no aesthetics scoring
- one JSON object only

### `agents/improver-instructions.md`

Improver guidance:

- read gaps
- make one targeted skill edit
- report what changed

## Directory Structure

```text
eval/
├── run-archetype-eval.mjs
├── run-eval.mjs
├── copilot-isolated-worker.mjs
├── lib/
│   ├── mode.mjs
│   ├── runtime/
│   ├── protocol/
│   ├── orchestrator/
│   ├── persistence/
│   └── reporting/
├── configs/
├── agents/
├── prompts/
├── results/
├── output/
├── criteria.md
├── patch-deps.cjs
└── package.json
```

## Practical Recommendations

### For stable baseline comparisons

- freeze `skills/shared/` for the duration of the run
- keep outputs in dedicated result folders
- prefer isolated or fresh-session runs
- set `mode: "baseline"` (or `--mode baseline` in `run-eval.mjs`)
- avoid improver edits during the comparison
- make transport, reasoning effort, and model the only changing variables

### For iterative skill tuning

- use `run-archetype-eval.mjs`
- keep prompts narrow to one archetype
- consider builder session reuse when you want prior review feedback to influence the next loop
- keep judge output JSON-only and inspect `judge_dimension_scores`, not just `totalScore`

### For debugging instability

- start with one config and low loop counts
- use isolated process mode first
- keep timeout values explicit
- inspect `builder_raw`, `judge_raw`, and artifact paths in saved results

## Summary

Today's eval framework already supports:

- explicit baseline vs tuning output separation
- Copilot SDK runtime selection with fresh, reused, and isolated execution modes
- strict request/response contracts plus typed failure categories
- persistence ledgers, manifests, artifact manifests, and generated run reports
- model-level reasoning effort tuning
- MCP transport experiments for builder runs
- isolated child-process execution
- session reuse across loops
- explicit JSON reviewer output

But the stronger architectural separation between **baseline model eval** and **skill tuning**, plus any ACP runtime work, is still a direction rather than a completed endpoint. Use the current mode split, run reports, and ledgers as the accurate source of truth for today's harness behavior.
