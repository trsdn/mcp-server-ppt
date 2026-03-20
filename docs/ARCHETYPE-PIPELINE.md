# Archetype Pipeline

This document explains how slide-design knowledge moves from curated or observed examples into the runtime design surface without leaking private source material.

## Two Knowledge Layers

PptMcp now combines two complementary layers:

### 1. Curated archetype guidance

Human-authored layout guidance lives under:

- `src\PptMcp.Core\Data\archetypes\`

These files describe:

- when to use a family
- layout rules
- density guidance
- anti-patterns
- representative variants

This layer is intentionally stable and editorial.

### 2. Learned reference coverage

Observed slide patterns come from curated example slides and are turned into a sanitized learned-reference catalog.

This layer captures:

- observed slide counts
- observed subtype counts
- example details
- audit samples

This layer is empirical and can grow as more good slides are curated.

## Privacy Rules

The reference pipeline is built around strict privacy constraints:

- intake must start from accepted slides only
- no screenshot galleries are committed under `src\`
- no customer names, source deck names, or provenance-bearing file names are exposed through public runtime outputs
- raw/private asset roots stay outside the committed runtime surface

The public runtime only exposes sanitized reference IDs and aggregated design knowledge.

## Canonical Intake Flow

### Step 1: Collect individual slides

Individual slides are extracted or copied into:

- `eval\input\individual-slides`

### Step 2: Direct LLM triage

The direct triage workflow inspects those slide images and separates them into:

- `eval\output\slide-triage\good`
- `eval\output\slide-triage\reject`

Only the accepted `good` set is allowed to feed the learned-reference pipeline.

### Step 3: Batch classification

Accepted slides are processed in batches.

The canonical provenance for each batch comes from:

- `eval\output\archetype-transfer\batch-lists\`

The corresponding classification output is captured as JSONL under:

- `eval\output\archetype-transfer\batch-results\`

Those batch lists are the source of truth for source-path normalization and replayability.

### Step 4: Normalize to the accepted taxonomy

Batch results are normalized so that:

- rogue or temporary labels are mapped back into the accepted top-level taxonomy
- subtype assignments become consistent across batches
- audit-only misbucketed examples can be tracked without silently changing the public taxonomy

### Step 5: Regenerate the sanitized catalog

The normalized batch outputs are then used to regenerate the sanitized learned-reference catalog.

At a high level, that regeneration step produces:

- manifest-style slide coverage
- subtype summaries
- learned-only family coverage
- sanitized example references safe for runtime exposure

## Where the Data Lives

### Public, committed guidance

- `src\PptMcp.Core\Data\archetypes\*.md`
- `src\PptMcp.Core\Data\archetypes\registry.md`

### Private or large eval assets

Preferred asset roots live under:

- `eval\input\`
- `eval\output\`
- `eval\results\`
- `eval\data\archetype-references\`

If you keep eval assets in a separate private repo clone, set:

- `PPTMCP_EVAL_ASSET_REPO_ROOT`

For a direct reference-catalog override, the runtime also supports:

- `PPTMCP_REFERENCE_DATA_ROOT`

## Runtime Surface

The runtime design surface is intentionally unified.

Use:

- `design list-archetypes`
- `design get-archetype`

to retrieve both:

- curated layout guidance
- learned observed coverage

This means a family can appear in the runtime catalog even if it is learned-first rather than fully authored as a curated layout file.

## Family Model

Today the pipeline distinguishes between:

- **17 curated families** with authored layout files under `src\PptMcp.Core\Data\archetypes\`
- additional **learned-only families** surfaced through the unified runtime catalog

Examples of learned-only or learned-first coverage include:

- `annotated-diagram`
- `chart-and-commentary`
- `org-chart`

Sub-archetypes capture the more specific observed shapes inside each top-level family.

## Refresh Checklist

When refreshing archetype/reference data:

1. Curate only from `eval\output\slide-triage\good`
2. Keep batch lists and batch results together
3. Normalize labels into the accepted taxonomy before runtime regeneration
4. Verify the regenerated catalog exposes only sanitized IDs and aggregated fields
5. Re-run the design CLI/MCP regression coverage after wiring new reference data
6. Update docs when the taxonomy or operating rules change

## Relationship to Other Docs

- [Eval Framework](../eval/README.md) explains how the harnesses, triage flows, and run artifacts work
- [Archetype Registry](../src/PptMcp.Core/Data/archetypes/registry.md) explains the curated family selection layer
- [Agent Client Architecture](AGENT-CLIENT.md) explains how runtime deck-building consumes this design knowledge
