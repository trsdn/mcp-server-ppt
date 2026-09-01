# Changelog

All notable changes to PptMcp (PowerPoint MCP Server) will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Security

- **Scriban bumped 6.6.0 → 7.2.6**: resolves NU1904 (critical) and four NU1902 (moderate) NuGet audit advisories that broke `dotnet restore` and caused the scheduled CodeQL workflow to fail on `main` since July
- **StreamJsonRpc bumped 2.24.84 → 2.25.29**: clears the transitive `MessagePack` 2.5.198 (2 high, 9 moderate) and `Nerdbank.MessagePack` 1.0.2 (1 high, 2 moderate) advisories. `dotnet list package --vulnerable --include-transitive` is now empty
- **All 33 open npm advisories resolved** (17 high, 14 moderate, 2 low): transitive dependencies were refreshed in `vscode-extension`, `src/PptMcp.Agent` and `eval`, clearing advisories in `undici`, `form-data`, `tmp`, `lodash`, `js-yaml`, `brace-expansion`, `picomatch`, `fast-uri`, `linkify-it`, `markdown-it`, `qs`, `uuid`, `@azure/identity` and `@azure/msal-node`. Only lock files changed — no declared dependency version was altered, and all affected packages are build-time only, so nothing shipped in the extension package was affected. `npm audit` now reports 0 vulnerabilities in all three manifests

### Fixed

- **11 MCP tools were advertised but every invocation failed with "Unknown command category"** (#124): `background`, `comment`, `customshow`, `headerfooter`, `pagesetup`, `placeholder`, `printoptions`, `shapealign`, `slideimport`, `smartart` and `tag` were listed by MCP discovery and exposed as CLI commands, but no call to any of them could ever reach the Core implementation. A third of the advertised tool surface was unreachable through both entry points.
  - ROOT CAUSE: MCP tools, CLI commands and per-category action dispatch are all source-generated from the `[ServiceCategory]` interfaces, but the top-level category switch in `PptMcpService` was hand-written. A migration to generated dispatch was started and left incomplete, so the switch covered 22 of 33 categories while discovery reported all 33. The generated `DispatchToCore` for the missing 11 already existed — only the wiring was absent, which is why the gap produced no compile error
  - FIX: `ServiceRegistryGenerator` now emits `ServiceRegistry.DispatchTable.g.cs`, a table with one entry per category built from the same interfaces the tools are generated from. `PptMcpService` looks routing up in that table instead of a hand-maintained switch, so a category added to Core is routable by construction. Action validation still happens before session acquisition, so an invalid action does not start PowerPoint
  - The 22 hand-written command fields and the `DispatchSimpleAsync`/`DispatchSessionless`/`TryParseDelegate` helpers they required are gone
- **`export extract-text` threw `RuntimeBinderException: Cannot convert type 'int' to 'bool'`** (#124): `Shape.HasTextFrame` returns `MsoTriState` (`msoTrue` is `-1`), not a VBA `Boolean`, so casting the boxed value to `bool` fails at runtime. 32 of 33 Core call sites already used `Convert.ToInt32(...) != 0`; `ExportCommands.ExtractText` was the lone outlier. `Chart.HasTitle` and `Chart.HasLegend` are genuine `Boolean` properties and are unaffected

### Changed

- **The MCP smoke test now proves tools are callable, not merely listed**: it previously asserted only that `ListToolsAsync` returned the expected 33 names, which is why #124 shipped undetected — a tool can be discoverable and unroutable at the same time. Every advertised tool is now driven through `PptMcpService` and must reject the *action* rather than the *category*. Added `ServiceRoutingTests`, which asserts the same invariant against the generated category list, and an end-to-end test that invokes a previously-unroutable category against real PowerPoint. None of the routing checks need PowerPoint, so the gate stays fast

- **The v1.1.0 release job went red although every package published** (#120): `Publish to Registries` failed with `Not on npm after publishing: ppt-cli-skill`, yet NuGet carried `PptMcp.McpServer` and `PptMcp.CLI` 1.1.0 and npm carried `ppt-mcp-skill` and `ppt-cli-skill` 1.1.0. A red release hides real failures, so the check itself was the defect.
  - ROOT CAUSE: the verification asked `npm view` right after publishing and treated npm's index lag as a failure. The window is widest for a package created in the same run — and both skill packages were created in that run. `npm view` also resolves through whatever registry the environment configured, so a proxy feed's 404 is indistinguishable from a genuinely absent package
  - FIX: the check now makes five attempts 15s apart (linear, because the lag is a short roughly constant index delay rather than congestion) and queries `https://registry.npmjs.org/<package>` directly, comparing the exact version string against the keys of the `versions` object. Each attempt logs what it asked and what it saw, so a failure says which versions npm did return
- **A version check could have matched the wrong version**: the Marketplace verification confirmed only that *some* listing existed, so a stale listing carrying the previous version would have been reported as a successful publish. It now requires the exact version, compared by string equality against each returned version rather than a substring or `-match` — `1.1.0` is a substring of `11.1.0`
- **npm bootstrap path removed**: `NPM_TOKEN` existed only to create a package that Trusted Publishing cannot create itself. Both skill packages now exist and have trusted publishers, so the token path, its preflight branch and the `NODE_AUTH_TOKEN` environment on both publish steps are gone and npm publishing is OIDC-only. The preflight now fails the release if a package it is asked to publish is missing from npm. The `NPM_TOKEN` repository secret is unused and can be deleted
- **The release job installed `npm@latest`, which is now npm 12**: npm 12 blocks dependency install scripts by default, and both places where the publish job upgrades npm asked for `@latest`. Pinned to `npm@^11.5.1` — above the OIDC minimum, below the breaking major (currently resolves to 11.19.0, where `@latest` resolves to 12.0.2). Neither skill package declares dependencies or lifecycle scripts, so this repository was not acutely exposed; the pin exists so a future npm major cannot break a release unannounced. Both sites carry a comment saying why, so it is not tidied back to `@latest`
- **A successful npm publish failed the release two seconds later**: during v1.1.0 both skill packages published correctly — the provenance attestations are in the Sigstore transparency log — but the verification step queried npm immediately and got a 404, which failed the `publish` job and skipped the GitHub release entirely.
  - ROOT CAUSE: npm's read path is eventually consistent, so a package is not queryable the instant `npm publish` returns. The check asked exactly once. The Marketplace verification had the same flaw, where it is worse: a freshly uploaded extension goes through validation first, so a single immediate query would fail every first-time publish
  - FIX: both verifications now poll — five minutes for npm, ten for the Marketplace — and report which target is still missing
- **npm's tokenless publishing would have silently failed after the switch to OIDC**: trusted publishing requires npm 11.5.1 or newer, and the pinned Node runtime bundles npm 10, which has no OIDC support and falls back to looking for a token instead of reporting the real problem. The publish job now upgrades npm and refuses to continue below 11.5.1, checked before the irreversible NuGet push and again at the point of use, because the second `setup-node` can put the bundled npm back in front

- **The v1.0.3 release reported four successful publishes but only three landed**: both npm skill packages never reached npm, and the VS Code Marketplace never received the extension, yet every step showed as successful in the job summary.
  - ROOT CAUSE: each of those steps carried `continue-on-error: true`, so a failed publish was indistinguishable from a successful one. Verified against the registries: `trsdn.ppt-mcp` is absent from the Marketplace because the `trsdn` publisher was never created, and `ppt-mcp-skill` and `ppt-cli-skill` are absent from npm because the workflow publishes with Trusted Publishing (OIDC), which cannot create a package that does not exist yet — the trusted publisher is configured per package in the npm UI, so a brand new package can never be bootstrapped that way. NuGet, the GitHub release assets and the MCP Registry (`io.github.trsdn/mcp-server-ppt`, live for 1.0.1-1.0.3) were unaffected
  - FIX: removed `continue-on-error` from the npm and MCP Registry publishes and added a post-publish npm listing check. npm publishing now accepts an optional `NPM_TOKEN` secret to bootstrap the packages once, and `.npmrc` is written explicitly so the tokenless OIDC path is not broken by an empty token line. The Marketplace publish is opt-in through a `publish_vscode` workflow input and is verified against the Marketplace API. Both targets are additionally checked in a preflight that runs *before* the NuGet push, since a version pushed to NuGet.org cannot be replaced
- **The MCP Registry job waited 20 minutes on every release for a URL that could never respond**: the NuGet propagation check requested `v3-flatcontainer/PptMcp.mcpserver/...`, but the flat container requires an all-lowercase package id, so every attempt 404'd and the loop always ran its full three attempts before falling through.
- **The shipped update check could never find an update**: both the CLI and the MCP Server queried `v3-flatcontainer/PptMcp.cli` and `v3-flatcontainer/PptMcp.mcpserver`, but the NuGet flat container API requires the package id lowercased (`LOWER_ID`) and returns 404 for any other casing. The 404 surfaced as an `HttpRequestException` that `NuGetVersionChecker` swallowed in a catch-all returning `null`, which is indistinguishable from "you are on the latest version" — so no user has ever been told that a newer release exists, on either entry point.
  - FIX: lowercased both ids. Added `check-nuget-flatcontainer-ids.ps1` as a pre-commit gate that scans source, workflows and documentation for mixed-case flat container ids, because this failure mode is a silent 404 rather than a crash and cannot be caught by running the code. The gate immediately found two documentation pages still describing the broken URL
- **Everything shipped pointed at a domain that does not exist**: `pptmcpserver.dev` (NXDOMAIN) was published as the `homepage` of both npm skill packages, the VS Code extension and the MCPB bundle, as the MCPB `documentation` and `privacy_policies` target, in the release notes, in the tray "About" dialog and across the skill and MCPB READMEs — 19 references in total. All now point at paths verified to exist in the repository
- **The MCPB bundle declared a privacy policy that was never written**: `privacy_policies` linked into the dead domain, and no privacy document existed anywhere in the repository. Added `docs/PRIVACY.md` describing what the tool does locally and documenting its single outbound request — the NuGet version check, which currently has no opt-out

- **The npm preflight checked that a token existed, not that it worked**: an invalid or wrongly scoped `NPM_TOKEN` would have passed the preflight and failed at the publish step — which runs *after* the NuGet push, so the version would already have been burned on NuGet.org, where it cannot be replaced. The preflight now authenticates against npm with `npm whoami` and reports npm's own rejection message
- **One unconfigured publishing target blocked every working one**: the npm preflight failed the entire release when the two skill packages did not exist and no `NPM_TOKEN` was set, so NuGet, the GitHub release assets and the MCP Registry — all of which work — could not be released at all. Added a `publish_npm` workflow input (default on, mirroring `publish_vscode`) that skips the npm path, and the release notes no longer advertise an npm install command when npm was not published
- **The Marketplace preflight would have blocked the first release with a false diagnosis**: it asked the extension query for the publisher's extensions and treated an empty list as "the publisher does not exist". A publisher that was just created has no extensions yet, which is precisely the state the preflight exists for, so the very first release after creating `trsdn` would have failed with an error telling the maintainer to create a publisher that was already there.
  - ROOT CAUSE: the extension query cannot distinguish the two states. Verified in both directions: `trsdn` (exists, empty) and a fabricated publisher name both return zero extensions, while `https://marketplace.visualstudio.com/publishers/<id>` returns 200 for `trsdn` and for `ms-vscode` and 404 for the fabricated name
  - FIX: the preflight now queries the publisher page and separates 404 from other failures, so an unreachable Marketplace no longer looks like a missing publisher. Both outcomes still stop the release, because the following step pushes to NuGet.org and cannot be undone
- **Marketplace publishing is documented as needing an Azure DevOps PAT that not every maintainer can obtain**: MFA requirements can make the PAT route inaccessible. Documented the officially supported manual route instead - the released VSIX already declares `"publisher": "trsdn"` and `"name": "ppt-mcp"`, so it can be uploaded on the publisher management page as `trsdn.ppt-mcp` with no Azure DevOps access at all
- **The npm bootstrap instructions could not be followed**: they asked for an "automation" token scoped to `ppt-mcp-skill` and `ppt-cli-skill`. Legacy automation tokens were removed in November 2025, and a granular token cannot be scoped to packages that do not exist yet - which is the entire point of the bootstrap. The instructions also omitted the **Bypass two-factor authentication** setting, which is off by default and without which the publish stops at a 2FA prompt no runner can answer. Corrected the steps, the workflow header comment and the troubleshooting entries

- **Pre-commit gates were reporting success without checking anything**: five of the ten gates were written against a hand-maintained `ToolActions.cs` / `ActionExtensions.cs` that the source generator architecture removed long ago.
  - ROOT CAUSE: `audit-core-coverage.ps1` parsed files that no longer exist, found zero methods, and printed "No gaps detected - 100% coverage maintained!" with exit code 0. It was also wired into two CI workflows with `-FailOnGaps`, so it produced a green required check while inspecting nothing. `check-mcp-core-implementations.ps1`, `check-cli-coverage.ps1`, `check-cli-action-coverage.ps1` and `audit-cli-actions.ps1` aborted on the same missing files
  - FIX: removed all five. CLI/MCP parity is now structural — both entry points are generated from the same Core interfaces — so the gate that was actually missing is a check that the *published* counts still match the generated surface. Added `check-documented-counts.ps1`, which derives the tool surface from the generated service registry (33 tools, 223 operations), verifies every count published across the Markdown documentation and the `FEATURES.md` table total, and fails if it detects nothing
- **`check-dynamic-casts.ps1` never ran**: the script contained an em dash inside a string literal. Windows PowerShell decodes these files as cp1252, which turned the character's trailing byte into a typographic quote and made the entire file a parse error, so the gate failed with a syntax error that looked like a policy violation.
  - FIX: made the script ASCII-only and added a per-file baseline (`scripts/dynamic-casts-baseline.txt`). The 140 pre-existing undocumented casts are tolerated so the hook is installable again, while any new undocumented cast fails the check
- **`check-cli-settings-usage.ps1` checked one command instead of three**: the gate silently skipped most of the CLI.
  - ROOT CAUSE: the Settings class was matched with a regex that ran to the end of the file, so unrelated JSON DTOs were reported as unused properties; the file filter `*Command.cs` did not match `*Commands.cs`, hiding `DiagCommands`, `ServiceCommands` and `SessionCommands`; and only the first Settings class per file was read, though `SessionCommands.cs` declares four
  - FIX: brace matching instead of a greedy regex, a corrected file filter, iteration over every Settings class, and a non-zero exit when no command is inspected
- **`pre-commit.ps1` swallowed gate failures**: the COM leak check, the SKILL.md auto-staging step and the dynamic cast check each caught their own errors and continued, so a gate that crashed was indistinguishable from a gate that passed. All three now abort the commit.
- **MCP smoke test gate failed on localized machines**: the guard that verifies the test filter actually matched something parsed the English `dotnet test` summary, so on a German Windows install it read `Bestanden! ... erfolgreich: 1` and reported "CRITICAL: No smoke tests passed" for a test that had just passed. The gate now pins `DOTNET_CLI_UI_LANGUAGE=en` for the duration of the run.

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
