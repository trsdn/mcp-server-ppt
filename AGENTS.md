# AGENTS.md

Guidance for AI coding agents working in this repository.

---

## 0. Repository identity — read this before anything else

**This repository is `trsdn/mcp-server-ppt`. Every operation targets it and only it.**

`trsdn/mcp-server-ppt` is registered on GitHub as a fork of `sbroenne/mcp-server-excel`, but that is metadata only. The two projects have **completely diverged** — they do not share a root commit, and upstream is unrelated spreadsheet automation code owned by someone else.

### Never

- ❌ Open, close, comment on, or modify **any** issue or pull request in `sbroenne/mcp-server-excel`
- ❌ Add an `upstream` remote, or any remote pointing at a repository you do not own
- ❌ Merge, rebase, or cherry-pick from upstream
- ❌ Fetch upstream tags into this clone
- ❌ Report upstream issues or PRs as if they belonged to this project

### Always

- ✅ Confirm the target repository before **any** `gh` command that reads or writes state
- ✅ Keep exactly one remote: `origin` → `https://github.com/trsdn/mcp-server-ppt.git`

### Verify

```powershell
git remote -v              # exactly one remote pair, both trsdn/mcp-server-ppt
gh repo set-default --view # trsdn/mcp-server-ppt
git tag                    # only this fork's tags (v1.0.0 - v1.0.3)
```

Repair if any check disagrees:

```powershell
git remote remove upstream
gh repo set-default trsdn/mcp-server-ppt
```

### Why this warning exists

`gh` derives its target from the git remotes. When an `upstream` remote is present it silently prefers the **parent** repository — so `gh issue list` and `gh pr list` return upstream's data with no indication that you are reading someone else's project.

This caused a real incident: an agent reported upstream issues #777 and #750 and PR #751 as belonging to this project, and was about to close issues in a repository it had no business touching. The correct answer was that this fork had zero open issues and zero open PRs.

The stale remote was also a live hazard — it carried a **push** URL, so a mistyped `git push upstream` would have written straight into another maintainer's repository.

**The fork relationship cannot currently be dissolved.** GitHub's "Leave fork network" requires the repository to have no child forks; this one has several. Detaching would also discard all stars, watchers, issues, PRs and child forks. So the fork status will persist — which is precisely why the remote and the `gh` default must stay pinned.

---

## 1. What this project is

**PptMcp** is a Windows-only toolset for programmatic PowerPoint automation via COM interop, aimed at coding agents and automation scripts.

There are **two equal entry points**, and every feature must work identically through both:

| Entry point | Path | Transport |
|---|---|---|
| MCP Server | `src/PptMcp.McpServer` | in-process |
| CLI | `src/PptMcp.CLI` | daemon over named pipe |

Supporting layers: `PptMcp.ComInterop` (COM patterns, STA threading, sessions), `PptMcp.Core` (PowerPoint business logic), `PptMcp.Service` (session management and routing), `PptMcp.Generators*` (source generators for CLI commands and MCP tools).

---

## 2. Non-negotiable rules

The authoritative list lives in [`.github/instructions/critical-rules.instructions.md`](.github/instructions/critical-rules.instructions.md). The ones that bite most often:

| Rule | Summary |
|---|---|
| 0 | Never commit without running the tests for what you changed |
| 1 | Never set `Success = true` alongside an `ErrorMessage` |
| 1b | Never wrap `batch.Execute()` in a catch that returns an error result — let exceptions propagate |
| 16 | Test **only** the feature you touched; the full suite takes 45+ minutes |
| 21 | **Never commit, push, or merge without explicit user approval** |
| 22 | COM cleanup belongs in `finally`, never in a swallowing `catch` |
| 24 | After changing a tool or action, sync **all** touchpoints (CLI, MCP, skills, READMEs, counts) |
| 26 | No confidential customer or project names in commits, PRs, or issues |
| 27 | Update `CHANGELOG.md` before merging |
| 29 | TDD: write the test first, watch it fail, then implement |
| 30 | Integration tests only — mocked COM tests prove nothing |
| 31 | Fork-only; see section 0 above |

---

## 3. Building and testing

```powershell
dotnet build PptMcp.sln -c Release      # must finish 0 warnings / 0 errors

# Surgical testing — pick the feature you changed
dotnet test tests\PptMcp.Core.Tests -c Release --no-build --filter "Feature=Slide&RunType!=OnDemand"

# End-to-end smoke test against real PowerPoint
.\scripts\Test-CliWorkflow.ps1

# Pre-commit gates
.\scripts\check-com-leaks.ps1
.\scripts\pre-commit.ps1
```

**If the build fails with `MSB3027` (file locked):** a `pptcli` daemon or MSBuild node still holds the output.

```powershell
Get-Process pptcli -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force }
dotnet build-server shutdown
Start-Sleep 4
```

Integration tests need PowerPoint installed and are slow (roughly 20-30 s per test). The first run after a rebuild is occasionally flaky from leftover daemon state — re-run once before believing a failure.

---

## 4. PowerPoint COM pitfalls

Hard-won specifics that are easy to get wrong:

- **`Presentation.SlideMasters` does not exist.** Masters are reached through `Presentation.Designs.Item(i).SlideMaster`, one master per design.
- **Layout names are localized**, and so is `CustomLayout.MatchingName`. There is no locale-independent layout identifier in the COM API. Resolve in stages: exact `Name`, then `MatchingName`, then position within the first design using the canonical Office layout order, which is identical across locales.
- **Indices are 1-based.** Use `Slides.Item(index)`.
- **Z-order changes require explicit reordering.**
- Before writing new COM code, look for a working example in another open-source project. [NetOffice](https://github.com/NetOfficeFw/NetOffice) is the best reference for essentially any PowerPoint COM API.

---

## 5. Release notes

Release tags on `origin` are `v1.0.0` through `v1.0.3`. **No release tag is an ancestor of `main`**, so `git describe --tags` fails in this repository.

Any tooling that derives the current version must sort semver tags directly rather than rely on ancestry — `git describe` silently falling back to `v0.0.0` would produce a version *below* what is already published. See `.github/workflows/release.yml`.
