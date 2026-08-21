# Changelog

All notable changes to PptMcp (PowerPoint MCP Server) will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Security

- **Scriban bumped 6.6.0 → 7.2.6**: resolves NU1904 (critical) and four NU1902 (moderate) NuGet audit advisories that broke `dotnet restore` and caused the scheduled CodeQL workflow to fail on `main` since July
- **StreamJsonRpc bumped 2.24.84 → 2.25.29**: clears the transitive `MessagePack` 2.5.198 (2 high, 9 moderate) and `Nerdbank.MessagePack` 1.0.2 (1 high, 2 moderate) advisories. `dotnet list package --vulnerable --include-transitive` is now empty
- **All 33 open npm advisories resolved** (17 high, 14 moderate, 2 low): transitive dependencies were refreshed in `vscode-extension`, `src/PptMcp.Agent` and `eval`, clearing advisories in `undici`, `form-data`, `tmp`, `lodash`, `js-yaml`, `brace-expansion`, `picomatch`, `fast-uri`, `linkify-it`, `markdown-it`, `qs`, `uuid`, `@azure/identity` and `@azure/msal-node`. Only lock files changed — no declared dependency version was altered, and all affected packages are build-time only, so nothing shipped in the extension package was affected. `npm audit` now reports 0 vulnerabilities in all three manifests

### Fixed

- **Pre-commit gates were reporting success without checking anything**: five of the ten gates were written against a hand-maintained `ToolActions.cs` / `ActionExtensions.cs` that the source generator architecture removed long ago.
  - ROOT CAUSE: `audit-core-coverage.ps1` parsed files that no longer exist, found zero methods, and printed "No gaps detected - 100% coverage maintained!" with exit code 0. It was also wired into two CI workflows with `-FailOnGaps`, so it produced a green required check while inspecting nothing. `check-mcp-core-implementations.ps1`, `check-cli-coverage.ps1`, `check-cli-action-coverage.ps1` and `audit-cli-actions.ps1` aborted on the same missing files
  - FIX: removed all five. CLI/MCP parity is now structural — both entry points are generated from the same Core interfaces — so the gate that was actually missing is a check that the *published* counts still match the generated surface. Added `check-documented-counts.ps1`, which derives the tool surface from the generated service registry (33 tools, 223 operations), verifies every count published across the Markdown documentation and the `FEATURES.md` table total, and fails if it detects nothing
- **`check-dynamic-casts.ps1` never ran**: the script contained an em dash inside a string literal. Windows PowerShell decodes these files as cp1252, which turned the character's trailing byte into a typographic quote and made the entire file a parse error, so the gate failed with a syntax error that looked like a policy violation.
  - FIX: made the script ASCII-only and added a per-file baseline (`scripts/dynamic-casts-baseline.txt`). The 140 pre-existing undocumented casts are tolerated so the hook is installable again, while any new undocumented cast fails the check
- **`check-cli-settings-usage.ps1` checked one command instead of three**: the gate silently skipped most of the CLI.
  - ROOT CAUSE: the Settings class was matched with a regex that ran to the end of the file, so unrelated JSON DTOs were reported as unused properties; the file filter `*Command.cs` did not match `*Commands.cs`, hiding `DiagCommands`, `ServiceCommands` and `SessionCommands`; and only the first Settings class per file was read, though `SessionCommands.cs` declares four
  - FIX: brace matching instead of a greedy regex, a corrected file filter, iteration over every Settings class, and a non-zero exit when no command is inspected
- **`pre-commit.ps1` swallowed gate failures**: the COM leak check, the SKILL.md auto-staging step and the dynamic cast check each caught their own errors and continued, so a gate that crashed was indistinguishable from a gate that passed. All three now abort the commit.

- **Required status checks could never pass on documentation- or dependency-only pull requests**: `build-cli` and `integration-runner-disabled` are required checks on `main`, but both workflows were restricted to source paths. A skipped check never reports a status, so any pull request that did not touch those paths stayed permanently blocked with "the base branch policy prohibits the merge".
  - ROOT CAUSE: a required status check combined with a `paths` filter — GitHub treats "skipped" and "never reported" identically for branch protection
  - FIX: removed the `paths` filter from the `pull_request` trigger of `build-cli.yml`, `integration-tests.yml` and `codeql.yml` so the required checks always report. The `push` filters are unchanged, so pushes to `main` still skip irrelevant builds
- **Dependency review license gate rejected the GitHub Copilot CLI**: `@github/copilot` and its nine platform-specific packages ship under their own `LICENSE.md` rather than an SPDX identifier, so the license checker classified them as incompatible and failed every pull request that touched those lock files. They are now exempt from the license check only — vulnerability scanning still applies to them.
- **`master` tool threw `RuntimeBinderException` on every action**: `list`, `list-shapes`, `edit-shape-text` and `list-layouts` all read `Presentation.SlideMasters`, which does not exist in the PowerPoint COM API.
  - ROOT CAUSE: property carried over from a spreadsheet-oriented code base; PowerPoint exposes masters through `Presentation.Designs.Item(i).SlideMaster` (one master per design)
  - FIX: added a shared `GetMaster(presentation, masterIndex)` helper that resolves masters through `Designs`, and routed all four actions through it
- **Layout lookup failed on non-English Office installations**: `slide create --layout Blank` and `slide apply-layout` raised "layout not found" because `CustomLayout.Name` and `CustomLayout.MatchingName` are both localized (for example `Leer`, `Titelfolie`, `Zwei Inhalte` on a German install).
  - ROOT CAUSE: lookup compared only against the localized `Name`, and the COM API exposes no locale-independent layout identifier
  - FIX: `FindLayout` now resolves in stages — exact `Name`, then `MatchingName`, then position within the first design using the canonical Office layout order (which is identical across locales). Canonical English names and numeric indices both work; the not-found message now lists the layouts that are actually available
- **Release workflow would have published a version below the current release**: the next release was calculated as `v0.1.0` even though `v1.0.3` is already published.
  - ROOT CAUSE: `git describe --tags` only finds tags that are ancestors of `HEAD`. No release tag is an ancestor of `main` in this repository, so the command failed and the `|| echo "v0.0.0"` fallback silently swallowed the error
  - FIX: the workflow now selects the highest semver tag with `git tag -l --sort=-v:refname`, which is independent of ancestry, and fails loudly if the tag cannot be parsed instead of emitting a bogus version
- **MCP resources advertised tools that do not exist**: `ppt://help/resources` and `ppt://help/quickref` told LLMs to call `powerquery`, `datamodel`, `namedrange`, `connection` and `range`, none of which are registered by this server, and documented `presentationPath`/`sessionId` parameters instead of the actual `path`/`session_id`.
  - ROOT CAUSE: the resource provider was carried over unchanged from the spreadsheet-oriented code base and never validated against the generated tool surface
  - FIX: both resources now describe the real tools (`slide`, `shape`, `text`, `notes`, `slidetable`, `comment`, `section`, `design`, `vba`, `export`, `file`) with the parameter names the generated tools actually expose
- **Agent skill CLI reference was never regenerated and shipped spreadsheet commands**: `skills/ppt-cli/references/cli-commands.md` listed `pivottable`, `slicer`, `powerquery` and `worksheetstyle`, and contained no parameters or actions at all.
  - ROOT CAUSE: three independent defects in `scripts/Build-AgentSkills.ps1` — it looked for `pptcli.exe` under a hardcoded, wrong target framework and silently skipped generation; it parsed the English section headers `COMMANDS:`/`OPTIONS:`, which Spectre.Console localizes, yielding zero commands on a non-English host; and it expected an `Actions:` prefix that the command descriptions do not contain
  - FIX: the target framework is now read from the csproj, section headers are matched locale-independently, actions are taken from the generated `ServiceRegistry` files, and every failure mode now throws instead of emitting an empty or stale reference

### Changed

- **Documentation corrected against the generated tool surface**: published tool and operation counts were wrong in every location that stated them (`FEATURES.md` claimed 204 operations in its header while its own table summed to 104 and omitted three tools; `mcpb/README.md` claimed 25 tools with 225 operations). All counts now state the verified **33 tools / 223 operations**, and `FEATURES.md` is generated from the tool surface rather than maintained by hand
- **Spreadsheet content removed from user-facing documentation**: tool lists, example prompts, workflows and prerequisites across the main, MCP server, CLI, VS Code extension, skill and package READMEs still described Power Query, DAX, PivotTables, ranges and slicers. `docs/SECURITY.md` documented a `--privacy-level` parameter and `docs/INSTALLATION.md` required an MSOLAP provider for DAX — neither exists in this code base (0 references in source)

### Added

- **`AGENTS.md`** at the repository root, documenting repository identity, the two equal entry points, build and test commands, and the PowerPoint COM pitfalls that are easy to get wrong
- **Rule 31 (repository identity)** in the critical rules: this project is `trsdn/mcp-server-ppt` and must never read from or write to its upstream. `gh` silently resolves to the parent repository when an `upstream` remote is present, which previously caused upstream issues to be misreported as belonging to this project

- Official source-side Copilot SDK agent client under `src\PptMcp.Agent`, including local planner tests and documentation for the agent architecture
- Dedicated documentation for the evaluation framework and the archetype/reference pipeline
- **33 PowerPoint MCP tools with 223 operations** for comprehensive PowerPoint automation via COM interop
- **Slide management** (7 ops) — list, read, create, duplicate, move, delete, apply-layout
- **Shape operations** (17 ops) — add, move, resize, fill, line, shadow, rotation, z-order, grouping, copy between slides, connectors, merge shapes (union/combine/fragment/intersect/subtract)
- **Text editing** (6 ops) — get/set text, find, replace, format (font, size, bold, italic, color, alignment)
- **Charts** (5 ops) — create, get info, set title, set type, delete
- **Slide Tables** (9 ops) — create, read, write cells, add/delete rows and columns, merge cells
- **Animations** (4 ops) — list, add, remove, clear effects
- **Transitions** (3 ops) — get, set, remove slide transitions
- **Design/Themes** (4 ops) — list designs, apply themes, get theme colors, list color schemes
- **Images** (1 op) — insert with position and size control
- **Speaker Notes** (3 ops) — get, set, clear
- **Sections** (4 ops) — list, add, rename, delete presentation sections
- **Hyperlinks** (4 ops) — add, read, remove, list
- **Slideshow** (4 ops) — start, stop, navigate, get status
- **Slide Masters** (1 op) — list masters and layouts
- **Export** (4 ops) — PDF, slide images (PNG), video (MP4), print
- **VBA Macros** (5 ops) — list, view, import, delete, run
- **Media** (3 ops) — insert audio/video, get media info
- **Window Management** (4 ops) — get info, minimize, restore, maximize
- **File Validation** (1 op) — test file accessibility
- **Document Properties** (2 ops) — get/set title, author, subject, etc.
- **Comments** (4 ops) — list, add, delete, clear slide comments
- **Placeholders** (2 ops) — list placeholders, set placeholder text
- **Slide Background** (3 ops) — get info, set solid color, reset to master
- **Headers & Footers** (2 ops) — get/set footer text, slide numbers, date
- **SmartArt** (2 ops) — get diagram info, add nodes
- **Shape Alignment** (2 ops) — align and distribute shapes on slides
- **Custom Shows** (3 ops) — list, create, delete custom slide shows
- **Page Setup** (2 ops) — get/set slide size and orientation
- **Slide Import** (1 op) — import slides from another .pptx file
- **Tags** (3 ops) — custom metadata on slides and shapes
- **MCP Server** — Model Context Protocol server for AI assistants (GitHub Copilot, Claude, ChatGPT)
- **CLI** (`pptcli`) — Command-line interface for scripting and coding agents
- **COM interop** — Uses PowerPoint's native COM API for 100% safe automation
- **Session management** — Shared sessions between MCP Server and CLI
- **Parameter validation** — All required string parameters validated before COM execution
- **COM resource safety** — All COM objects released in finally blocks to prevent leaks
