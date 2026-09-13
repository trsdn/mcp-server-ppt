# Contribution and Retirement

Assessment date: 2026-09-10.
Latest contribution update: 2026-09-13.

Goal: retire `trsdn/mcp-server-ppt` as an independently maintained product and contribute selected improvements to [sbroenne/mcp-server-powerpoint](https://github.com/sbroenne/mcp-server-powerpoint).

Working list: [Feature Contribution Checklist](#feature-contribution-checklist). This is the current inventory to select work from; the earlier numbered PR proposals below remain historical context.

## Recommendation

Contribute a small sequence of independently useful fixes, then selected authoring capabilities. Do not propose a repository merge, a replacement architecture, or full feature parity. Preserve the old repository as an attributed reference, not a second product that must keep pace.

The reverse interview established:

- Stefan welcomes contributions, according to Torsten.
- Full feature parity is not a prerequisite for archiving.
- The preferred sequence is balanced: correctness fixes first, features next.

Start by offering PRs 1 and 2. Coordinate PR 3 against Stefan's existing startup-cleanup issue. Offer PRs 5 and 6 as the first feature batch. PRs 4 and 7 are valuable but require explicit agreement on their contracts before implementation.

**Stefan merged #73, theme palettes #76, and composition recipes #77 on September 11.** Audit #74 remains open with its review fixes published as `53bd305` and all remote checks passing. Stefan also merged #75 with startup cleanup; the local startup experiment remains parked. Selected-shape text find/replace is published as [PR #82](https://github.com/sbroenne/mcp-server-powerpoint/pull/82), authored by `trsdn`. Remaining proposals are not prerequisites for retirement. Earlier PR bodies and submission snapshots below are retained as planning history.

## Evidence and Limits

Local source was inspected at commit `d9ec6ab69971b19f94b8b92239755fbbcff0687f`; the worktree was clean before this document was added. Stefan's public `main` was inspected through source files, interfaces, documentation, and issue pages. Its repository page reported head `acf92b9736a68d8c18b8fb3a8158f9fcbf627cd4`. Reads were not an atomic checkout; recheck the target revision and issue status before implementation.

The local session store initially contained zero sessions. Reindexing found one session but zero turns, files, or references. Searches for the repository name, Stefan, contributions, comparison, and archiving recovered no earlier investigation. No matching investigation document was found in workspace Markdown either. This does not prove that the earlier conversation never existed, only that it was unavailable through this local index.

Verification performed:

- Read the target Core test project: it lacks an explicit `IsTestProject` marker, consistent with issue #71.
- Read the target COM leak gate and evaluated its exact two matching predicates against an in-memory example containing two acquisitions and one release. It reported clean.
- Read the target startup/cleanup implementation, shared fixture, public command interfaces, feature audit, and skill packaging script.
- Compared those against local implementations, tests, guidance, and the changelog. Historical measurements in the changelog are not new measurements from this investigation.
- Initial research did not run a target build or integration test. Subsequent implementation evidence is recorded below. Startup cleanup remains an issue-backed investigation, not a newly established root cause.
- Initial research made no remote changes. Subsequent authorized submission is recorded below. The original checkout's remote remains unchanged. No archiving was performed.

Evidence labels used below: **Confirmed** means inspected source or a local static probe; **Reported** means the maintainer's issue describes a reproduction; **Proposed** means a design recommendation requiring agreement.

## Implementation Evidence

### September 13 Update

- **Audit #74:** review fixes published as `53bd305` through SSH authenticated as
	`trsdn`. Covers loop/using declarations, captured assignments, wrapped borrowed
	receivers, and exact supported cleanup-helper names. The first expanded test run
	reproduced 12 failures; final scanner tests passed **47/47**. Unchanged pre-commit
	passed with Release **0 warnings/errors**, MCP **41/41**, packaging **112/112**.
	The real-source scan found 95 acquisitions, with zero unmatched releases, but is
	explicitly not a proof of leak freedom or `finally` placement. Both review replies
	were posted as `trsdn`; PR remains open. Remote checks: five successful, one
	changeset check skipped, zero failing or pending.
- **Selected-shape find/replace:** isolated branch `feature/text-find-replace` in
	`../mcp-powerpoint-text-pr`, based on target `main` at `bee71dc`. Adds literal
	`textframe find-text` and `replace-text` through the shared Core contract, with
	case/whole-word options, 1-based original match positions, deletion through an
	explicit empty replacement, and no searching of inserted text.
- Initial real-COM regressions failed on each missing operation before implementation.
	An MCP round-trip exposed whitespace searches being rejected by generated validation;
	the existing parameter attribute corrected it without changing generators.
	Full TextFrame domain **56/56**, focused MCP **2/2**, and a live isolated CLI
	find/deletion/whitespace workflow passed. Formatting outside a replacement and
	another shape's text were verified after save/reopen.
- Release build **0 warnings/errors**; strict site build, **36-page** audit, and
	**23-source** deploy coverage passed. Unchanged full pre-commit gate passed:
	TextFrame **56/56**, full MCP **44/44**, packaging **77/77**, plus source audits
	and documentation counts (**16 tools / 189 operations**).
- Published commit `829207f` to the `trsdn` fork and opened **PR #82** against
	Stefan's `main`. Author and remote head verified; feature worktree is clean.
	Initial CI checks are queued/running, with no inline review findings yet.
	No merge or archive was performed.

Text evidence: `%TEMP%/text-domain-ad7fc016ca304905b2c1aab1ead1addf`,
`%TEMP%/text-mcp-green-6ba4c33afb14459595f758fcad2d42c9`,
`%TEMP%/text-build-7b3d469ff28b477993247ccd35426765`, and
`%TEMP%/text-pr-gates-05a8204286f44c0b90e4f9b326443036`.

### Earlier Feature Publication Snapshot

Following explicit approval to create valid PRs, both features were isolated into
fresh worktrees based directly on current target `main` (`090c8b1`, after #75),
tested, committed, pushed to `trsdn/mcp-server-powerpoint`, and submitted against
Stefan's `main`. Neither depends on the other or includes the parked startup
candidate or the audit/discovery PR changes. The original dirty checkout is preserved.

- [Theme palettes #76](https://github.com/sbroenne/mcp-server-powerpoint/pull/76):
	branch `feature/theme-color-palettes`, implementation `2350863`, review follow-up
	`b68041b`; worktree `../mcp-powerpoint-theme-pr`. Isolated Master **16/16** and
	schema **1/1** passed without a discovery override. The unchanged pre-commit
	gate passed: Release **0 warnings/errors**, Master **16/16**, MCP **42/42**,
	packaging **65/65**. Strict site build, **35-page** audit and **22-source**
	deployment coverage passed. The site check exposed a stale exact-action list,
	corrected before submission. Initial remote CI was entirely green.
- Copilot's two #76 review comments identified three additional stale count
	claims: bundle manifest, agent instructions, and native-feature audit. All were
	corrected in `b68041b`, with replies posted. Follow-up pre-commit passed with
	protocol **37/37** and packaging **65/65**; the commit message's **32/32** was
	a transcription error, corrected in the public review reply. CI reruns for the
	follow-up are pending at this recording.
- [Composition recipes #77](https://github.com/sbroenne/mcp-server-powerpoint/pull/77):
	branch `feature/composition-recipes`, commit `db43cd4`; worktree
	`../mcp-powerpoint-recipes-pr`. Isolated recipe tests **2/2** passed at both
	aspect ratios, without a discovery override. Unchanged pre-commit gate passed:
	Release **0 warnings/errors**, MCP **41/41**, packaging **65/65**. The recipe
	class was run separately because the gate selects only modified Core domains.
	Strict site build, **36-page** audit and **23-source** deployment coverage passed.
	CI Gate, Docs Site, changeset and dependency checks are green; CodeQL is still
	running and no automated review findings have appeared at this recording.

At this earlier snapshot both PRs were open and mergeable, with minor changesets.
No merges, archiving, force-pushes, remote changes, or updates to #73/#74 were
performed during that publication step. The September 13 update above supersedes
these PR states; Stefan subsequently merged #73, #76, and #77.

Publication evidence: `%TEMP%/palette-pr-gates-637289ea9ade4ee2800205ef443623f8`,
`%TEMP%/palette-review-gates-737bc06ee26c44728d93ee3416fa32c0`,
`%TEMP%/recipes-pr-gates-68cceec8e6d14e8eb7405ca4b62355b7`, and recipe TRX/PNGs
under `%TEMP%/recipes-pr-tests-5ccfb0264d534c9284b49304544ff49d`.

Earlier implementation evidence below was collected before branch isolation;
the publication evidence above supersedes its checkout/discovery limitations.

- **Theme palettes:** `master get-theme-colors` reads twelve named `#RRGGBB`
	colors for a selected 1-based master. Typed PIA implementation, generated CLI
	and MCP surfaces, canonical skill guidance, synchronized public counts, and a
	changeset are included. Existing native BGR font/background results are unchanged.
- Five new real-COM cases cover role names, channel order, distinct design
	palettes, save/reopen persistence, and invalid indices. The initial test failed
	to compile before the API existed. Final affected master group: **16/16 passed**.
- New MCP tool-schema check: **1/1 passed**. A live CLI run on a unique daemon
	pipe created a presentation, returned twelve colors with `masterIndex`,
	`masterName`, and `themeColors`, then closed the session and requested daemon
	shutdown successfully.
- **Composition recipes:** comparison, chart and insight, three-step timeline,
	and metric callout. Shared guidance includes geometry, content limits, failure
	cases, and a reproducible test command; deck-builder links to it. Both skill
	reference trees contain the new guide. All example content is synthetic.
- Real PowerPoint render tests passed separately at **720x540 and 960x540**
	(**2/2 cases**), checking text read-back/bounds, native chart category count,
	and PNG export. All eight exported images were visually inspected: no visible
	text overlap or clipping. This is example validation, not a measured improvement
	in LLM-generated decks or a general automatic layout guarantee.
- Final Release solution build: **0 warnings, 0 errors**. Existing shared-guide
	plugin packaging test: **1/1 passed**. Public count gate: **16 tools / 187
	operations**. Core interface completeness, success-flag, COM audit, and diff
	whitespace checks passed. The COM audit does not prove typed PIA ownership.
- Core runs use `-p:IsTestProject=true` until the separate discovery fix lands.
	Every feature test runner had a ten-minute limit and required a positive TRX
	executed count; none timed out. No old-repository tests supply this evidence.

Local evidence: `%TEMP%/pptmcp-feature-checks-8c5eb4de8c1d4d93921bff562aa5738e`
(master/protocol/packaging TRX),
`%TEMP%/pptmcp-palette-cli-41f6982fdd5740d7a1042c09d9ac71be` (CLI JSON),
`%TEMP%/pptmcp-recipes-43-24652dc3cb2b443e85856f1e1816c195/images/720x540`, and
`%TEMP%/pptmcp-recipes-wide-125b088e18b64fd786510a120516581b/images/960x540`.

### PR 1 Verification and Submission

PR 1 is submitted as [sbroenne/mcp-server-powerpoint#73](https://github.com/sbroenne/mcp-server-powerpoint/pull/73), following explicit approval to commit, push, and open the PR. It was implemented and verified in the sibling checkout `../mcp-server-powerpoint`, on branch `fix/core-test-discovery`, based on `acf92b9736a68d8c18b8fb3a8158f9fcbf627cd4`. Contribution commit: `a63a38005946136dc632c67532ec30594cb1157c`. Only the Core test project and its metadata regression test changed.

- Added `CoreTestProject_IsExplicitlyMarkedForTestDiscovery` to the target's existing metadata-test class. It failed with expected `true`, actual `null` before the fix: 1 executed, 1 failed.
- Added `<IsTestProject>true</IsTestProject>` to the target Core test project. The identical regression command then passed: 1 executed, 1 passed, 0 skipped.
- Built `Sbroenne.PowerPointMcp.slnx` in Release successfully with SDK 10.0.400; no warnings or errors were reported. The SDK was installed locally without changing the repository's pin.
- Restore initially failed because of direct NuGet TLS connectivity. The user-supplied package proxy resolved restore through a command-local source override. No machine-specific feed was added to either repository.
- The first `Feature=Chart` integration run discovered Core tests but was canceled after approximately 153 seconds, before the configured ten-minute limit. No completed TRX was produced; no PowerPoint or test-host process remained at the subsequent check. This is not passing runtime evidence.
- The single retry completed successfully in approximately 240 seconds: **22 executed, 22 passed, 0 failed, 0 skipped**. It used the same filter, a fresh TRX directory, and a ten-minute runner-tree limit. No `IsTestProject` command-line override was used.
- The structured result is retained locally under `%TEMP%/pptmcp-chart-08151bdf4d4a4fccaac8942b3411db85/chart.trx`. Runner exit code was 0. The subsequent process check found no PowerPoint or test-host process remaining; no force-kill was needed on the successful run.
- Environment: .NET SDK 10.0.400, test runtime 10.0.11, Office Click-to-Run 16.0.20501.20002.
- The unchanged pre-commit script passed: Release build with 0 warnings/errors, MCP Server tests 41/41, packaging tests 66/66, and all configured static gates. Neither test suite skipped any cases. The script did not select Core tests because no Core domain source changed; the separate Chart run above supplies that evidence.
- Created the `trsdn/mcp-server-powerpoint` fork and pushed only the contribution branch there. Both local checkouts' remotes were left unchanged. The PR targets Stefan's `main` and carries `skip-changelog` under the test-only changeset policy.
- At the first post-submission check, CI was queued/running and there were no inline review comments or submitted reviews. Remote CI success and maintainer approval are not yet established.

The successful runtime command, invoked within the ten-minute runner wrapper from the target checkout, was:

```powershell
& "$env:LOCALAPPDATA\PptMcpContribution\dotnet\dotnet.exe" test tests\PowerPointMcp.Core.Tests -c Release --no-build --no-restore --filter 'Feature=Chart' --logger 'trx;LogFileName=chart.trx' --results-directory $resultDirectory --nologo
```

`$resultDirectory` was a newly generated temporary directory for this invocation. The wrapper required a fresh TRX with a positive executed count and a zero runner exit code.

The metadata regression, solution build, scoped Core runtime acceptance, and pre-commit gates passed before committing and submitting. This does not establish the root cause of startup-cleanup issue #72 or claim that the entire Core suite passed. No merge or archive operation has been performed.

### PR 2 Verification and Submission

PR 2 is implemented in the target checkout on the independent branch
`fix/com-acquisition-audit`, starting from the same target `main` base above rather
than PR 1's branch. It is submitted as
[sbroenne/mcp-server-powerpoint#74](https://github.com/sbroenne/mcp-server-powerpoint/pull/74),
with contribution commit `85e3fd35f268d3443718b1f4926d69a2fee20ac3`.
At submission, PRs 3-7 were proposals; the current feature status is recorded above.

- Replaced whole-file matching in `scripts/check-com-leaks.ps1` with Roslyn syntax
	matching using PowerShell's bundled parser. Local dynamic acquisitions require a
	matching release for the same variable in the owning function and lexical scope.
- Borrowed aliases are excluded, including existing-variable dispatch aliases and
	direct `ctx`/`context` aliases of `Presentation` and `App`. The first actual-source
	scan exposed seven such aliases; these were inspected and the classification was
	corrected with failing-then-passing regressions, not a suppression baseline.
- Added 27 executable scanner cases to the existing metadata/tooling test suite.
	They cover unrelated releases, neighboring scopes, deferred assignment, aliases,
	comment/string decoys, escaped identifiers, nullable-type trivia, generated files,
	and missing, empty, or malformed input. The original unrelated-release regression
	failed against the old scanner before implementation. Later alias and syntax
	corrections also have observed red/green evidence.
- Final focused run: **27 passed, 0 failed, 0 skipped**, approximately 23 seconds.
	Final Release solution build succeeded without reported warnings or errors.
- Final actual-source scan: 111 source files, 4 explicit session ownership-file
	exemptions, 93 dynamic acquisition variables, 0 unmatched releases. This is not a
	claim of typed PIA coverage or absence of all COM leaks.
- Updated the target site's contributor guide to document supported patterns and
	limitations. Typed/`var` ownership, inactive preprocessor branches, fields,
	ownership transfers, and control-flow/release-in-finally verification remain out
	of scope. No runtime source or package references changed.
- IDE diagnostics and `git diff --check` are clean. The strict documentation build,
	35-page site audit, and deploy-path coverage for 22 mirrored sources passed after
	using CI's offline star-history asset generation step. Dependencies were installed
	in the target's ignored virtual environment through the user-supplied PyPI feed.
- The user explicitly approved running the remaining gates and then committing,
	pushing, and submitting PR 2 if they pass. The unchanged pre-commit gate passed
	under a 20-minute runner-tree limit: Release build with 0 warnings/errors, all
	static audits, **MCP 41/41** in 6 minutes 50 seconds, and **packaging 92/92** in
	1 minute 32 seconds. Neither suite skipped any tests; no timeout kill was needed.
- Pushed only the contribution branch to `trsdn/mcp-server-powerpoint` and opened
	the PR against Stefan's `main`. Remotes remain unchanged. The PR carries the
	`skip-changelog` label for internal tooling and documentation changes.
- At the initial post-submission check, CI was still running and no submitted
	reviews or inline review comments were present. Remote CI success and maintainer
	approval are not yet established. No merge or archive operation was performed.

Focused target test command:

```powershell
dotnet test tests\PowerPointMcp.SkillGeneration.Tests -c Release --no-build --no-restore --filter 'FullyQualifiedName~ComLeakAudit_' --nologo
```

## Compatibility Before Porting

### PR 3 Investigation Status

On 2026-09-10, both submitted PRs (#73 and #74) had passing CI, with no inline
review findings. Issue #72 remained open with no comments; target `main` remained
at `acf92b9736a68d8c18b8fb3a8158f9fcbf627cd4`.

A local real-COM test now reproduces a related partial-creation failure: hold an
exclusive lock on a unique destination, invoke the shared creation helper through
`ReopenPresentation`, and inspect the application's remaining presentation count.
The original code threw from `Presentation.SaveAs` and left one presentation where
zero were expected. This is consistent with a failed-save ownership gap, but later
runs showed that collection counts alone do not identify which presentation was
left behind. It does not establish the startup timeout cause reported in #72.

The experimental cleanup is **not passing**. Directly closing the acquired
presentation also throws COM error `0x80004005` after the failed save. The original
save error is preserved, with the close error attached as `PresentationCleanupError`.
The latest single-test run executed one test and failed, with a fresh TRX under
`%TEMP%/pptmcp-startup-close-0f004ad42d6c434491e72f4a185ab006/startup.trx`.
Every run used a ten-minute runner-tree limit. A process check after the earlier
diagnostic run found no remaining PowerPoint process.

Only the target creation helper and the existing save-test file have local,
uncommitted experimental changes. They are currently on `fix/com-acquisition-audit`
but are **not included in published #74**. No new branch, commit, push, or PR was
made for this investigation. Further cleanup diagnostics are required before this
can be presented as a fix.

#### Follow-Up on 2026-09-11

The user authorized continued local investigation, not publication. Setting the
failed new presentation's `Saved` flag before closing produced one passing `.pptx`
run (1 executed, 1 passed), but **did not establish a reliable fix**. A strengthened
test for `.pptm` still reported `Close()` error `0x80004005` and one remaining
presentation. A subsequent `.pptx` case failed before creation because
`ReopenPresentation` retained a released `_presentation` after clearing `_context`;
its next call dereferenced the null context.

A diagnostic attempted to distinguish a later STA callback from releasing the
destination lock. Its latest run did not reach those probes: two presentations
already existed before the forced failure, `Close()` reported no cleanup error,
and two remained where the count-based test expected one. The shared reopen path
suppresses errors while closing the old presentation, so aggregate count changes
cannot by themselves prove that the failed new presentation was closed. Replace
this assertion with identity-based tracking or an isolated constructor failure test
before treating a green result as acceptance evidence.

Evidence retained in fresh temporary directories:

- `pptmcp-startup-saved-21f60ebbc31e4fd086212e2d0db54942/startup.trx`: 1 passed;
	this is the initial, weaker count-only test, not final acceptance.
- `pptmcp-startup-recovery-832451b1812f4bbca12f4dd131281735/startup.trx`: 2 failed
	(cleanup failure followed by null-context recovery failure).
- `pptmcp-startup-count-b35048223e204d4e9f07f99b3991ae48/startup.trx`: 1 executed,
	1 failed; `.pptm` close failure with one presentation remaining.
- `pptmcp-startup-retry-0dcde969c8a14781be753eab7154f9bb/startup.trx`: 1 executed,
	1 failed; baseline count two, remaining count two, no reported cleanup error.

All paths above are under `%TEMP%`. Runs used SDK 10.0.400, a ten-minute runner-tree
limit, and `-p:IsTestProject=true` because this local branch does not include PR 1.
The typed `Saved` probe initially failed compilation because the referenced
PowerPoint PIA requires the absent Office assembly for `MsoTriState`; the subsequent
documented late-bound probe compiled. No timeout kill was needed, and no PowerPoint
process was present at the final process check.

At that point, the two target files contained **uncommitted, failing experimental code**, including
the saved-flag and delayed-close diagnostics. No new branch, commit,
push, PR, merge, or archive action was taken. Normal shutdown policy is unchanged.

#### Isolated Startup Baseline on 2026-09-11

With user approval, the shared-fixture regression and speculative production cleanup
were removed. Both files are restored to their committed content. The target now
has only one untracked addition:
`tests/PowerPointMcp.ComInterop.Tests/PresentationStartupTests.cs`.

The replacement drives `PresentationBatch` construction directly with an exclusively
locked, unique destination. It subscribes to the existing owned-process identity
event and retains a handle to that exact PID/start-time match before creation fails.
It requires the original `SaveAs` COM error, bounded process exit, and exit code 0.
It neither uses aggregate presentation counts nor silently accepts a force-killed
process. Shutdown diagnostics are captured through the existing logger parameter.
The lock file is deleted on disposal; the event subscription and process handle
are released in `finally`. No shared fixture or additional PowerPoint session is
created by the test.

Results against **unchanged production code**:

- Initial `.pptx` baseline: 1 executed, 1 passed, approximately 13 seconds.
	Evidence: `%TEMP%/pptmcp-startup-isolated-6b31e7aea8f844e3bbad855714ea8f9d/startup.trx`.
- Expanded `.pptx`/`.pptm` run: **2 executed, 1 passed, 1 failed, 0 skipped**.
	`.pptm` passed; `.pptx` logged `Application.Quit()` failure `0x80004005` and
	exited with code **-1**, consistent with the constructor's existing force-kill
	fallback. The failed case lasted approximately 17 seconds.
	Evidence: `%TEMP%/pptmcp-startup-isolated-formats-a7bde9c368f04d41b4be8d7cc51bba92/startup.trx`.

This establishes an intermittent failure to exit cleanly after a forced startup
save error. It does **not** reproduce the 120-second startup timeout or persistent
window reported in #72. The earlier passing run is baseline evidence, not evidence
of a fix. There is no production fix in the current diff.

The ComInterop test project already enables test discovery, so these runs did not
need `-p:IsTestProject=true`. Both runs were bounded by the same ten-minute runner
limit, which did not fire. Unlike earlier diagnostics, the second run detected an
application-level abnormal exit despite no wrapper timeout. No PowerPoint process
remained at the final check. Build output reported no warnings/errors and the new
test has no IDE diagnostics. No commit, push, or PR was created.

#### Early Startup Ownership Candidate

Continued local work registers the acquired presentation with `RunStaThread`'s
existing `startupPresentation` cleanup variable immediately after `Presentations.Add`
or `Open`, instead of waiting for the entire helper (including `SaveAs`) to return.
A private optional acquisition callback connects the helper to that owner. The
existing shared reopen caller is unchanged. No new COM operation, retry policy,
saved-flag mutation, force-kill policy, or public API was introduced.

The original ownership gap is visible in source: a thrown `SaveAs` previously
bypassed the assignment to `startupPresentation`, leaving outer cleanup with only
the application reference. The candidate lets the existing shutdown service attempt
close/release of that partial presentation before quitting the application.

Focused evidence with the candidate:

- Previously failing isolated `.pptx` case: **1 executed, 1 passed**, approximately
	12 seconds. `%TEMP%/pptmcp-startup-owner-c63a739d0f3f4c6cbc9b536329c4b46c/startup.trx`.
- Both formats in sequence: **2 executed, 2 passed, 0 skipped**, approximately
	14 seconds. `%TEMP%/pptmcp-startup-owner-formats-351c81163bf04623a0f753c39e04af5a/startup.trx`.
	Each retained process handle reported exit code 0; tests still require the
	original `SaveAs` COM error.
- Existing `SaveCopyAs_CreatesCopyAndLeavesSessionPathUnchanged` compatibility
	workflow: **1 executed, 1 passed**, covering successful creation, save-copy, and
	reopening through the unchanged shared-fixture caller.
	`%TEMP%/pptmcp-startup-compat-77bd0e63414441e990c5ad851340dbc7/compat.trx`.
	This Core test used `-p:IsTestProject=true` because PR 1 is not on this branch;
	the isolated ComInterop tests did not require that override.
- Full Release solution build succeeded with no reported warnings/errors. The
	first invocation used the wrong working directory and selected the original
	repository's SDK pin; rerunning from the target checkout resolved it without
	configuration changes.
- Target COM acquisition audit passed: 111 source files, 4 explicit session-file
	exemptions, 93 dynamic acquisitions, 0 unmatched releases. This gate does not
	independently validate the modified session file or typed ownership.
- IDE diagnostics and tracked-diff whitespace checks are clean. The ten-minute
	runner limits did not fire.

**Status: local candidate with passing focused checks, not ready for publication.**
The baseline failure was intermittent, so these passes do not establish complete
reliability. The reported #72 startup timeout and persistent-window scenario remain
unreproduced. Full pre-commit gates, release metadata, and a separate contribution
branch remain pending before any authorized publication. Current target changes
are the production helper and the new isolated startup test, both uncommitted on
`fix/com-acquisition-audit`; published #74 is unchanged.

The repositories share the same architectural family: COM/STA infrastructure, Core commands, a service layer, generated CLI/MCP surfaces, and persistent sessions. That is useful conceptual overlap, not proof that commits can be cherry-picked cleanly. No common-ancestor comparison was performed.

Important differences:

| Area | Contribution implication |
| --- | --- |
| COM access | Stefan requires typed PowerPoint PIA members wherever available. Much of this repository uses `dynamic`; port behavior and tests, not entire command classes. |
| Lifecycle | This repository deliberately terminates PowerPoint on normal shutdown. Stefan explicitly forbids force-killing a healthy process on the normal path. Do not port the local shutdown design or promise its teardown timings. |
| Test isolation | Stefan reuses a PowerPoint process through a shared fixture while creating fresh presentations. Preserve that convention instead of moving to one process per test. |
| Public contracts | Stefan uses `presentation`, `textframe`, `table`, and 1-based shape indices. This repository uses different names and often shape names. Keep Stefan's contracts. |
| Domain organization | The target advertises 16 tools/186 operations; this repository advertises 33/224. These are documented counts, not an apples-to-apples capability score. Several target tools consolidate domains split here. |
| Packaging and skills | Stefan already has both CLI/MCP skill packages and deterministic shared-guidance synchronization. Do not propose those as missing features. |
| Safety policy | Physical printing and window chrome are explicitly excluded by Stefan's native-feature audit. Local availability is not a reason to propose them. |
| Licensing | Both repositories use MIT licensing. Preserve applicable attribution for substantial reused code or text. Sanitized identifiers alone do not establish redistribution rights for source slides. |

## Ranked PR Sequence

Sizes describe review/implementation scope, not elapsed-time promises.

| PR | Proposed title | Priority | Size | Reuse and readiness |
| --- | --- | --- | --- | --- |
| 1 | Fix Core test discovery under the current .NET SDK | First | XS | Confirmed target defect; directly addresses #71. |
| 2 | Make the COM acquisition audit reject unrelated releases | First | M | Confirmed shared defect; adapt the local gate's principle, not its limitations. |
| 3 | Clean up partial presentation startup without changing normal shutdown | First, coordinate | M/L | Maintainer-reported #72; local failure/ownership tests provide useful scenarios. Root cause still needs reproduction. |
| 4 | Fail test gates when no tests execute and bound stalled runs | Next | M | Local runner provides a starting point; remove unsafe cleanup assumptions. |
| 5 | Expose the active presentation's theme color palettes | First feature | M | Local implementation and color regression tests; target interface gap confirmed. |
| 6 | Add focused slide-composition recipes to shared authoring guidance | First feature | S/M | Curated local content; no new tool/runtime required. |
| 7 | Add bounded text search and replacement for a selected text frame | Later feature | M | Target gap confirmed; local implementation must be rewritten, not copied. |

PR 1 enables trustworthy Core-test execution for subsequent runtime PRs. PR 2 can be developed independently. PR 3 must not wait for an elaborate test-runner redesign. PR 4 can follow PR 1, but must not hide PR 3's leak by cleaning it silently. PRs 5 and 6 are independent. PR 7 does not require a catalog, agent, or evaluation framework.

Keep one behavioral outcome per PR. Include that outcome's tests, relevant skill/XML documentation, generated CLI/MCP parity, and the target repository's release metadata in the same PR. Do not put all tests into a later omnibus PR. Do not mechanically import this repository's minimum test/doc counts or pre-push policy into Stefan's project.

## PR 1: Restore Core Test Discovery

**Proposed title:** `fix(tests): enable Core integration test discovery`

**Evidence:** Confirmed source gap and maintainer-reported reproduction in [issue #71](https://github.com/sbroenne/mcp-server-powerpoint/issues/71). The [target test project](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/tests/PowerPointMcp.Core.Tests/PowerPointMcp.Core.Tests.csproj) omits `IsTestProject`. The issue reports that the ordinary command exits successfully without running tests, while forcing the property discovers 22 Chart cases.

### Proposed PR Body

The documented Core integration-test command can currently succeed without discovering or executing any tests. The pre-commit hook invokes the same project and trusts the exit code, so a successful gate does not establish that the modified Core behavior was exercised.

This PR will explicitly identify the Core project as a test project. It will verify the documented feature-filtered command without a command-line property override and add a narrowly scoped regression check that prevents the marker from disappearing unnoticed.

This is a test configuration fix. It will not change chart behavior, replace the shared fixture, alter PowerPoint shutdown, or introduce a new repository-wide testing policy.

### Scope and Acceptance

- Given the normal project command without `-p:IsTestProject=true`, a real existing feature filter discovers and executes at least one case.
- Verify the same invocation shape used by the pre-commit hook. Capture the actual discovered/executed count; do not hardcode the issue's historical 22 cases as a permanent invariant.
- A regression check rejects a test project with the required marker removed, using the target's existing project/metadata test conventions.
- A resulting startup timeout is reported as a real test failure, not misclassified as a discovery regression or solved by undoing the marker.

**Dependencies:** None for the marker fix; #72 may prevent a clean full execution result on an affected host. Report that honestly. Do not disable tests again to obtain green output.

**Why separate:** This is the smallest accepted-value contribution and unblocks credible evidence for every later runtime change. General zero-test protection belongs in PR 4.

## PR 2: Make the COM Leak Gate Discriminating

**Proposed title:** `fix(tooling): match COM acquisitions to their releases`

**Evidence:** Confirmed in the [target scanner](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/scripts/check-com-leaks.ps1). It asks whether a file has any `dynamic` declaration and any `ComUtilities.Release` text. Those predicates also return clean when one variable is released and another is leaked.

**Local contribution:** [scripts/check-com-leaks.ps1](../scripts/check-com-leaks.ps1), plus the distinction between acquired objects and borrowed context aliases. The local implementation is still dynamic-only and matches names across a file; it is not a complete ownership or control-flow analyzer.

### Proposed PR Body

The current audit can accept a real leak whenever the same file contains an unrelated release. It can also print a successful result after inspecting no relevant acquisitions. This makes the gate weaker than its output suggests.

This PR will replace whole-file existence checks with acquisition-specific checks for the supported late-bound COM patterns. A release must correspond to the acquisition in the same owning scope. Borrowed aliases of context-owned objects must remain exempt so the checker does not encourage callers to release the presentation they do not own.

The gate will explicitly state what it verifies: a matching release for supported acquisitions. It will not claim that a regex proves release-in-finally, exception safety, typed PIA ownership, or the absence of all COM leaks. Typed coverage will either be implemented and tested using the target's existing parsing infrastructure or remain an explicitly documented follow-up; a renamed copy of the local dynamic-only scanner is not typed coverage.

### Scope and Acceptance

- A file with two acquired variables and a release for only one fails and names the unreleased variable.
- Reusing a variable name in a neighboring method does not let that method's release discharge this method's acquisition.
- Declare-null-then-acquire patterns are recognized; borrowed context aliases do not produce findings.
- Comments, string literals, and generated output do not count as real acquisitions/releases.
- Broken source discovery fails. A legitimate zero-late-bound-acquisition result must be explicitly distinguished from a broken scan in a PIA-first codebase, not universally rejected.
- Deliberate red/green probes demonstrate each supported rule. The output identifies unsupported typed/control-flow cases without calling the whole repository leak-free.

**Dependencies:** No PowerPoint session is required to exercise scanner rules. Real command repairs uncovered by the gate need their own focused COM tests. Use the target's existing tooling-test conventions.

**Why separate:** Establish a trustworthy signal before broad cleanup. Do not attach an unreviewable sweep through every command to the scanner PR. Baseline reviewed pre-existing findings only if necessary; never silently auto-accept new findings.

## PR 3: Fix Partial Startup Cleanup

**Proposed title:** `fix(com): clean up failed presentation startup safely`

**Evidence:** [Issue #72](https://github.com/sbroenne/mcp-server-powerpoint/issues/72) reports a 120-second fixture startup timeout followed by a responsive PowerPoint process remaining open. The [shared fixture](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/tests/PowerPointMcp.Core.Tests/SharedPresentationFixture.cs) obtains its batch in its constructor, so it cannot dispose a batch whose constructor never returned.

**Local contribution:** Failure and foreign-process scenarios from [tests/PptMcp.ComInterop.Tests/Integration/Session/PptBatchTimeoutTests.cs](../tests/PptMcp.ComInterop.Tests/Integration/Session/PptBatchTimeoutTests.cs) and [tests/PptMcp.ComInterop.Tests/Integration/Session/ProcessOwnershipTests.cs](../tests/PptMcp.ComInterop.Tests/Integration/Session/ProcessOwnershipTests.cs). These are test ideas, not an approved shutdown implementation.

### Proposed PR Body

A session constructor must either return an owned, usable batch or clean up the resources it acquired before failing. A failed shared-fixture constructor cannot rely on the fixture's later `Dispose` method to repair partial startup.

This PR will reproduce the failure at the session-startup boundary and fix the demonstrated ownership gap there. It will preserve normal graceful shutdown and the shared-fixture reuse model. It will not add routine process termination, lengthen timeouts to conceal the failure, or attribute a pre-test startup error to Chart commands.

The current implementation already has constructor and STA-finally cleanup. The investigation must determine why those paths do not reclaim the reported instance. One concrete path to examine is `OpenOrCreatePresentationCom`: it acquires a presentation and performs slide creation and Save As before returning it, while the outer startup owner records the presentation only after the helper returns. A failure or stall between acquisition and that handoff is a specific hypothesis, not yet the established cause of #72.

### Scope and Acceptance

- Reproduce and record the failing startup phase before changing behavior. Log acquisition, ownership handoff, cancellation, and cleanup outcomes without private paths.
- A failure after acquiring a presentation but before publishing the batch leaves no untracked owned presentation/COM resources.
- The documented repeated-run scenario succeeds after a failed startup, rather than inheriting stranded state.
- A pre-existing user presentation remains open with unsaved content intact. A test must never begin by killing every PowerPoint process to make this assertion easier.
- Normal successful close remains graceful. Emergency cleanup requires evidence of ownership; PID plus start time identifies a process but does not by itself prove the session created it.
- Teardown failure cannot replace the original startup error or silently report successful cleanup.

**Dependencies:** PR 1 for reliable Core test execution. Agree a bounded failure-injection/reproduction strategy with Stefan before coding if the failure is intermittent.

**Why separate:** Lifecycle work has substantially higher risk than the other proposals. Keep diagnostics and the demonstrated repair together; do not combine it with test-runner policy or the local fast-shutdown design.

## PR 4: Require Test Execution Evidence

**Proposed title:** `fix(tooling): reject empty test runs and bound stalled execution`

**Evidence:** The [target pre-commit hook](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/scripts/pre-commit.ps1) judges `dotnet test` by exit status. #71 demonstrates why that is insufficient. Existing instructions require timeouts but the hook's direct test calls do not themselves establish a timebox.

**Local contribution:** [scripts/Invoke-GuardedTest.ps1](../scripts/Invoke-GuardedTest.ps1). Reuse the principles of positive execution evidence, explicit filtering, and runner-tree cleanup. Do not copy its broad daemon termination, PID-snapshot-only ownership inference, output parsing, or timing assumptions wholesale.

### Proposed PR Body

Fixing Core test discovery addresses the present defect, but an empty filter or a future discovery regression can still return exit code zero. The gate needs positive evidence that the requested tests actually ran.

This PR will introduce a small shared test invocation wrapper for the existing gate. It will inspect a fresh per-invocation structured test result, reject zero executed cases, preserve nonzero test failures, and enforce a configurable wall-clock timeout. On timeout it will terminate the test runner and its descendants so an orphaned test host cannot continue executing tests after the gate has returned.

PowerPoint cleanup is a separate ownership problem. The wrapper will not kill a process merely because its PID first appeared during the test interval. A user can open PowerPoint during the same interval. Unproven ownership must result in a diagnostic and a failed/incomplete cleanup result, not destructive guessing.

### Scope and Acceptance

- An intentionally nonexistent filter exits nonzero even when `dotnet test` returns zero.
- Missing, malformed, or stale result files cannot satisfy the run. A suite where every case is skipped does not count as executed coverage.
- A genuine passing run succeeds; failing cases and nonzero runner exit codes remain failures.
- A deliberately timed-out test-runner process tree stops, including its test host, and the wrapper returns a distinct timeout failure.
- Both directory and project-file invocation forms work; paths containing spaces preserve argument boundaries.
- Existing non-COM protocol tests are not accidentally classified as PowerPoint-dependent. Do not make all development runs launch Office.
- Cleanup does not conceal leaks from PR 3: an owned survivor is recorded as failure even if emergency cleanup subsequently succeeds.

**Dependencies:** PR 1; agreement on structured-result format and timeout defaults. Keep this within the existing hook/CI policy, not a new mandatory full-suite pre-push manifest system.

**Why separate:** This is reusable prevention, not the immediate discovery repair. If runner lifecycle dominates the diff, split positive test evidence from timeout/tree cleanup into two PRs; do not delay the evidence fix.

## PR 5: Read Theme Color Palettes

**Proposed title:** `feat(master): expose theme color palettes for template-aware authoring`

**Evidence:** The target [presentation interface](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/src/PowerPointMcp.Core/Presentation/IPresentationCommands.cs) reads a theme name, while its [master interface](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/src/PowerPointMcp.Core/Master/IMasterCommands.cs) reads placeholder fonts/backgrounds and lists masters. Neither exposes the twelve-role theme palette.

**Local contribution:** `GetColors`, `ListColorSchemes`, and the shared reader in [src/PptMcp.Core/Commands/Design/DesignCommands.cs](../src/PptMcp.Core/Commands/Design/DesignCommands.cs); [tests/PptMcp.Core.Tests/Integration/DesignColorRoundTripTests.cs](../tests/PptMcp.Core.Tests/Integration/DesignColorRoundTripTests.cs); OLE color conversion in [src/PptMcp.ComInterop/ComUtilities.cs](../src/PptMcp.ComInterop/ComUtilities.cs).

### Proposed PR Body

Agents can apply a template and inspect its theme name, but cannot retrieve its semantic color palette. As a result, new charts and shapes may use guessed colors rather than the presentation's own accents.

This PR will add read-only theme-palette inspection using the existing master domain, subject to agreement on the selector. It will enumerate modern designs and their themes, return the twelve named roles, and identify the design/master each palette belongs to. It will not use the legacy `Presentation.ColorSchemes` collection, which the local project found unsuitable for modern OOXML theme enumeration.

The port will use typed PIA access wherever available, with a narrowly documented late-bound Office-theme boundary only where the restored interop metadata requires it. It will reuse the target's color conventions or add an explicitly named RGB/hex representation without silently changing existing color outputs.

### Scope and Acceptance

- A known theme returns exactly `Dark1`, `Light1`, `Dark2`, `Light2`, `Accent1` through `Accent6`, `Hyperlink`, and `FollowedHyperlink`.
- An asymmetric fixture color such as `#0B3D91` is read correctly, not as `#913D0B`. Verify against independently configured raw COM/theme data, not only a symmetric helper round trip.
- A multi-design presentation returns/selects each design's actual palette. The selector agrees with the target's master enumeration; never assume unrelated collection indices are interchangeable.
- Invalid selection returns the target's established validation response and does not mutate the presentation.
- A real COM read failure does not produce a successful incomplete palette.
- Generated CLI/MCP signatures, result serialization, and authoring guidance agree. Mixed read/write tool hints must remain honest at the aggregate tool level.

**Dependencies:** PR 1 and a maintainer-approved action/selector contract. Prefer one useful palette query over a new top-level design tool and duplicate list/get APIs. No dependency on the archetype catalog.

**Why separate:** Theme fonts are related but have different semantics: Latin/East Asian/complex-script roles, missing entries, and resolved placeholder fonts are not the same thing. Add font-scheme inspection in a later PR if requested.

## PR 6: Curated Slide-Composition Guidance

**Proposed title:** `docs(skills): add focused composition recipes with rendered examples`

**Evidence:** The target's [deck-builder guide](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/skills/shared/deck-builder.md) already contains planning, broad composition choices, and export-to-verify. The contribution is more precise recipe content, not introducing those existing workflows.

**Local contribution:** The authored family guidance under [src/PptMcp.Core/Data/archetypes/registry.md](../src/PptMcp.Core/Data/archetypes/registry.md), with provenance boundaries described in [docs/ARCHETYPE-PIPELINE.md](ARCHETYPE-PIPELINE.md).

### Proposed PR Body

The current skill explains how to assemble a deck and verify the rendered result. This PR will add a small set of actionable composition recipes that make those instructions more specific for recurring slide types.

The initial scope is four families: comparison, chart-with-insight, timeline/roadmap, and headline metric. Each recipe will describe content fit, information hierarchy, layout proportions, density limits, failure cases, and a reproducible example using only existing target operations. Positioning will derive from the actual slide dimensions, not assume every deck is 960 by 540 points.

The material will live in canonical shared guidance and use the existing packaging pipeline for both skills. There will be no new design tool, runtime catalog loader, private reference corpus, Copilot SDK dependency, or automatic prompt-tuning loop.

### Scope and Acceptance

- Each recipe has a distinct content purpose and an example rendered by Stefan's server, with a repository-owned or synthetic source deck.
- Every command/action referenced exists on the target surface. Translate local `text`, `slidetable`, shape-name, and session syntax; do not ship examples that invoke this repository's APIs.
- Verify representative 16:9 and 4:3 layouts using exported images. Text, labels, and diagrams stay within intended bounds at the declared density.
- The existing skill package tests confirm both distributions receive the canonical guidance.
- Keep examples small and load recipes on demand. Do not dump the entire local taxonomy into every model prompt.
- Review reused text and assets for authorship, licensing, and private provenance. Do not include learned examples merely because their filenames were sanitized.

**Dependencies:** Stefan agrees that these four recipes fit the skill's voice and intended audience. No runtime dependency on PR 5; recipes can use an explicitly declared example palette until theme inspection exists.

**Why separate:** This offers distinctive value at low runtime risk. First establish that the guidance is useful; a queryable catalog can be a later proposal if static, on-demand references are insufficient.

## PR 7: Bounded Text Search and Replacement

**Proposed title:** `feat(textframe): add bounded find and replace operations`

**Evidence:** The [target textframe interface](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/src/PowerPointMcp.Core/TextFrame/ITextFrameCommands.cs) supports whole-frame text and formatting but no find/replace action. Local `Find`/`Replace` live in [src/PptMcp.Core/Commands/Text/TextCommands.cs](../src/PptMcp.Core/Commands/Text/TextCommands.cs).

### Proposed PR Body

Updating a recurring term currently requires callers to retrieve a frame's text and rewrite it. A native text-range operation can locate and replace matches without treating the complete frame as one new string.

This PR will add find and replace for one explicitly selected text frame, following the existing 1-based slide and shape indices. Results will contain structured match positions/counts rather than only a human-readable message. Search options and replacement behavior will be explicit, and replacement will visit a finite set of original matches rather than repeatedly processing text it just inserted.

The local implementation is not suitable for direct reuse: it restarts `Find(searchText)` after each replacement, which can keep matching when the replacement equals or contains the search term. It also overwrites the found-range proxy without deterministic release. The contribution is the workflow and regression cases, implemented afresh against the target's conventions.

### Proposed Contract and Acceptance

- Initial scope is a selected shape's text frame. No presentation-wide traversal, table cells, SmartArt, grouped-child addressing, speaker notes, regex language, or cloning in this PR.
- Literal, non-overlapping matching; case sensitivity and whole-word behavior are explicit. Proposed defaults are case-insensitive and not whole-word-only, consistent with the native API; Stefan confirms them before implementation.
- Reject an empty search term. Permit an empty replacement to delete matches. A shape without a text frame produces the target's ordinary validation error.
- Replacement equal to the search term terminates. Replacement containing the search term terminates without revisiting inserted text. Adjacent matches and replacements of different lengths have exact counts.
- `find` does not mutate. Match offsets are explicitly 1-based in the selected COM text range. Record the indexing convention so callers do not assume .NET string indexing for all Unicode cases.
- Mixed formatting outside matched spans and existing unrelated hyperlinks remain unchanged. Agree and test the formatting applied inside replacement spans against real PowerPoint behavior.
- Unexpected COM failures surface. Do not promise transactional rollback across multiple replacements unless it is actually implemented; document that a failed mutation may be partial.
- Every acquired match range is released; save/reopen preserves the intended text; MCP and CLI return equivalent results.

**Dependencies:** PR 1 and contract agreement. Microsoft's [TextRange.Replace documentation](https://learn.microsoft.com/office/vba/api/powerpoint.textrange.replace) provides the native search continuation parameter; verify the restored PIA signature and runtime behavior before choosing the implementation.

**Why separate:** One selected frame is a manageable first contribution. Add deck-wide traversal or clone-with-replace only after match semantics and formatting preservation are reliable. Run-level formatting is a different user operation and deserves its own contract.

## Feature Contribution Checklist

Compared on 2026-09-13 against target `origin/main` at `bee71dc`, after fetching
`sbroenne/mcp-server-powerpoint`. This inventories exposed source capabilities,
not fresh runtime validation or a promise of drop-in compatibility. Equivalent
actions under different tool names are excluded. Refresh the target before starting
each item, since another contribution may already have closed the gap.

Each unchecked item is a candidate, not approved implementation work. Stable `FC`
IDs are local tracking IDs, not GitHub issue numbers. All 36 candidate issues were
created as `trsdn` on 2026-09-13 in `trsdn/mcp-server-ppt`; see the
[Contribution Issue Index](#contribution-issue-index) for the mapping. They are
proposed, not started. Group related operations into one bounded contribution; split
an item further if its contracts or risks differ. None is an archive prerequisite.

### First Candidates

- [ ] **FC01 - Contribute shape alignment and distribution.** Align selected shapes and distribute spacing; define slide-relative versus selection-relative behavior.
- [ ] **FC02 - Contribute theme-font inspection.** Read major/minor font roles for a selected master; define script coverage and preserve unresolved entries. Read-only.
- [ ] **FC03 - Contribute slide visibility controls.** Hide/unhide slides and control display of inherited master shapes.
- [ ] **FC04 - Contribute named custom shows.** Create, list, and delete named subsets of slides; validate order and slide references.

Source references: [shape alignment](../src/PptMcp.Core/Commands/ShapeAlign/IShapeAlignCommands.cs), [theme inspection](../src/PptMcp.Core/Commands/Design/IDesignCommands.cs), [slide controls](../src/PptMcp.Core/Commands/Slide/ISlideCommands.cs), [custom shows](../src/PptMcp.Core/Commands/CustomShow/ICustomShowCommands.cs).

### Authoring Candidates

- [ ] **FC05 - Contribute shape duplication and cross-slide copying.** Preserve editable shapes and define naming, placement, and clipboard implications.
- [ ] **FC06 - Contribute shape formatting transfer.** Copy formatting between existing shapes without replacing content.
- [ ] **FC07 - Contribute shape merging.** Define supported merge operations and which input shapes are consumed.
- [ ] **FC08 - Contribute shape scaling and aspect-ratio locking.** Scaling factors and explicit lock state; preserve existing absolute-size operations.
- [ ] **FC09 - Contribute shape gradient fills and opacity.** These are shape fills, not the slide/master gradients already available in the target.
- [ ] **FC10 - Contribute WordArt and 3D rotation.** Separate if needed; target bevel support already exists and is not a gap.
- [ ] **FC11 - Contribute attached connectors.** Bind connector endpoints to shapes, beyond the target's coordinate-based connector creation.
- [ ] **FC12 - Contribute text-frame layout controls.** Margins, word wrapping, and vertical alignment; auto-size already exists in the target.
- [ ] **FC13 - Contribute advanced text formatting.** Strikethrough, subscript, superscript, paragraph/character spacing, numbered bullets, and indentation. Split typography and paragraph controls if needed.
- [ ] **FC14 - Contribute substring hyperlinks.** Apply links to matching text rather than to the entire shape; define repeated-match handling.
- [ ] **FC15 - Contribute text insertion and case utilities.** Change case; insert symbols, date/time, and slide-number fields into shape text.
- [ ] **FC16 - Contribute slide cloning with replacement.** Define clone count, replacement scope, formatting behavior, and partial-failure semantics; keep separate from FC30.
- [ ] **FC17 - Contribute slide naming.** Set slide names without conflating them with titles or indices.
- [ ] **FC18 - Contribute master-shape text access.** List master shapes and edit their text, beyond title/body placeholder font properties.
- [ ] **FC19 - Contribute picture and inherited backgrounds.** Set picture backgrounds and restore inheritance without covering existing solid/gradient actions again.
- [ ] **FC20 - Contribute table-cell text formatting.** Boldness, font size, and alignment; cell fill and borders already exist in the target.
- [ ] **FC21 - Contribute chart-type changes.** Change an existing chart's type; define behavior for incompatible data layouts.
- [ ] **FC22 - Contribute picture transparent-color control.** Distinct from the target's grayscale/black-and-white/watermark recolor modes.
- [ ] **FC23 - Contribute SmartArt layout, style, and hierarchy changes.** Existing node creation/deletion is not a gap.
- [ ] **FC24 - Contribute animation inspection and editing.** Effect details, editing existing timing, reordering, and clearing sequences; basic add/delete already exists.
- [ ] **FC25 - Contribute media playback settings.** Volume, mute, and fades; media insertion and inspection already exist.

Source references: [shape](../src/PptMcp.Core/Commands/Shape/IShapeCommands.cs), [text](../src/PptMcp.Core/Commands/Text/ITextCommands.cs), [master](../src/PptMcp.Core/Commands/Master/IMasterCommands.cs), [background](../src/PptMcp.Core/Commands/Background/IBackgroundCommands.cs), [table](../src/PptMcp.Core/Commands/SlideTable/ISlideTableCommands.cs), [chart](../src/PptMcp.Core/Commands/Chart/IChartCommands.cs), [image](../src/PptMcp.Core/Commands/Image/IImageCommands.cs), [SmartArt](../src/PptMcp.Core/Commands/SmartArt/ISmartArtCommands.cs), [animation](../src/PptMcp.Core/Commands/Animation/IAnimationCommands.cs), [media](../src/PptMcp.Core/Commands/Media/IMediaCommands.cs).

### Inspection and Workflow Candidates

- [ ] **FC26 - Contribute slide/shape inspection.** Dedicated listings, detailed reads, slide summaries, and shape filtering by type. Compare target resources as well as tools before defining the gap.
- [ ] **FC27 - Contribute text diagnostics.** Word counts and a dedicated empty-placeholder audit; do not duplicate existing accessibility checks.
- [ ] **FC28 - Contribute hyperlink inventory and validation.** Define what validation proves, including internal destinations versus external addresses.
- [ ] **FC29 - Contribute proofing controls.** Text-language read/write and spelling-check invocation. Investigate modal/UI behavior before proposing unattended spelling checks.
- [ ] **FC30 - Contribute slide/deck-wide find and replace.** Extend beyond one selected text frame only after PR #82's contract is settled; explicitly define traversal and partial edits.
- [ ] **FC31 - Contribute video export.** Define completion, timeout, codec, and output-file verification behavior.
- [ ] **FC32 - Contribute text and image extraction.** Separate output formats, traversal, and image-fidelity expectations from rendered slide export.
- [ ] **FC33 - Contribute table/row convenience operations.** Whole-table and row reads, plus row writes; these aggregate existing cell capabilities rather than add new native features.
- [ ] **FC34 - Contribute notes convenience operations.** Append, clear, and read all notes; assess value over composing existing get/set calls.
- [ ] **FC35 - Contribute transition convenience operations.** Remove transitions and copy settings to all slides; basic get/set already exists.
- [ ] **FC36 - Contribute cross-session slide copying.** Different from the target's file-based import; define source/destination session ownership and design preservation.

Source references: [text](../src/PptMcp.Core/Commands/Text/ITextCommands.cs), [hyperlinks](../src/PptMcp.Core/Commands/Hyperlink/IHyperlinkCommands.cs), [proofing](../src/PptMcp.Core/Commands/Proofing/IProofingCommands.cs), [export](../src/PptMcp.Core/Commands/Export/IExportCommands.cs), [notes](../src/PptMcp.Core/Commands/Notes/INotesCommands.cs), [transitions](../src/PptMcp.Core/Commands/Transition/ITransitionCommands.cs), [slides](../src/PptMcp.Core/Commands/Slide/ISlideCommands.cs).

### Pending and Already Delivered

- **Pending:** selected-frame find/replace, [target PR #82](https://github.com/sbroenne/mcp-server-powerpoint/pull/82), head `829207f`. Do not create a duplicate contribution issue; FC30 is the broader follow-up.
- **Merged:** theme colors, [target PR #76](https://github.com/sbroenne/mcp-server-powerpoint/pull/76), and composition recipes, [target PR #77](https://github.com/sbroenne/mcp-server-powerpoint/pull/77).
- **Not gaps:** comments, sections, tags, placeholders, basic shape hyperlinks, footer/date settings, Save As/Copy, linked pictures, and accessibility reading order.
- **Not an existing local feature:** general run-level formatting. Substring hyperlinks exist here, but ordinary formatting still targets the whole shape; do not describe range formatting as a ready-to-port contribution.

### Deferred or Excluded

These are inventory items, not ordinary contribution tickets. Obtain scope agreement
before promoting one; do not file implementation issues for explicitly rejected capabilities.

- **FD01 - Design catalogs.** Archetypes, palettes, style profiles, grids, density profiles, deck sequences, and slide patterns. Curate a small subset and verify provenance rather than port the catalog wholesale.
- **FD02 - Agent client.** Copilot SDK plan/execute/verify/repair orchestration; substantial new runtime and maintenance responsibility.
- **FD03 - Visual evaluation pipeline.** Requires runner, model-access, cost, data-rights, and evidence-quality agreement; existing recipes do not imply acceptance of this system.
- **FD04 - VBA.** Module management, macro execution, and trust checks require a separate security/scope decision.
- **FD05 - Live slideshow control.** Start/stop, navigation, status, and configuration require a separate unattended-operation review.
- **FD06 - Persisted print settings.** Deferred in the target pending demonstrated demand; not permission to print.
- **Excluded - Physical printing and window controls.** Target policy explicitly rejects physical printing and desktop zoom/view/window controls. Record the difference, but do not propose them as routine ports.

Sources: [design contract](../src/PptMcp.Core/Commands/Design/IDesignCommands.cs), [agent client](AGENT-CLIENT.md), [evaluation](../eval/README.md), [VBA](../src/PptMcp.Core/Commands/Vba/IVbaCommands.cs), [slideshow](../src/PptMcp.Core/Commands/Slideshow/ISlideshowCommands.cs), [print options](../src/PptMcp.Core/Commands/PrintOptions/IPrintOptionsCommands.cs).

### Issue Tracking Procedure

Issue destination: **`trsdn/mcp-server-ppt`**, not Stefan's repository.
The user approved the full FC01-FC36 batch on 2026-09-13. Existing issues were
checked for duplicates before publication. The closed local theme-font issue #71
tracks the original implementation, not FC02's contribution to the target.
For future additions, obtain publication approval and check for duplicates first.

Use each checklist item's title with its ID, for example:
`[Contribution FC01] Shape alignment and distribution`.

Every issue should record:

- The user workflow and bounded proposed target behavior.
- Local source references and the target revision checked for an equivalent feature.
- Scope exclusions, unresolved API questions, and dependencies on existing PRs.
- Acceptance checks: real-COM regression first, positive executed test counts, both CLI/MCP entry points, persistence where relevant, and target-required gates.
- Status: proposed, accepted, implementing, PR open, merged, deferred, or declined.
- Links to the local checklist, any split issues, and the eventual target PR.

After filing, add the actual issue link to the corresponding FC issue-index entry. Check an item
only after its target contribution is merged; record a declined item explicitly rather
than checking it as delivered. Missing source capabilities alone are not evidence that
the maintainer wants a feature or that the local implementation is production-ready.

### Contribution Issue Index

All issues below are proposed contributions, opened by `trsdn`. No target PRs were
created as part of this batch. FD01-FD06 and explicitly excluded capabilities were
not filed. The source checklist remains local until separately committed/published;
each issue includes its scope and source references so it can be used independently.

- FC01: [#185 - Shape alignment and distribution](https://github.com/trsdn/mcp-server-ppt/issues/185)
- FC02: [#186 - Theme-font inspection](https://github.com/trsdn/mcp-server-ppt/issues/186)
- FC03: [#187 - Slide visibility controls](https://github.com/trsdn/mcp-server-ppt/issues/187)
- FC04: [#188 - Named custom shows](https://github.com/trsdn/mcp-server-ppt/issues/188)
- FC05: [#189 - Shape duplication and cross-slide copying](https://github.com/trsdn/mcp-server-ppt/issues/189)
- FC06: [#190 - Shape formatting transfer](https://github.com/trsdn/mcp-server-ppt/issues/190)
- FC07: [#191 - Shape merging](https://github.com/trsdn/mcp-server-ppt/issues/191)
- FC08: [#192 - Shape scaling and aspect-ratio locking](https://github.com/trsdn/mcp-server-ppt/issues/192)
- FC09: [#193 - Shape gradient fills and opacity](https://github.com/trsdn/mcp-server-ppt/issues/193)
- FC10: [#194 - WordArt and 3D rotation](https://github.com/trsdn/mcp-server-ppt/issues/194)
- FC11: [#195 - Attached connectors](https://github.com/trsdn/mcp-server-ppt/issues/195)
- FC12: [#196 - Text-frame layout controls](https://github.com/trsdn/mcp-server-ppt/issues/196)
- FC13: [#197 - Advanced text formatting](https://github.com/trsdn/mcp-server-ppt/issues/197)
- FC14: [#198 - Substring hyperlinks](https://github.com/trsdn/mcp-server-ppt/issues/198)
- FC15: [#199 - Text insertion and case utilities](https://github.com/trsdn/mcp-server-ppt/issues/199)
- FC16: [#200 - Slide cloning with replacement](https://github.com/trsdn/mcp-server-ppt/issues/200)
- FC17: [#201 - Slide naming](https://github.com/trsdn/mcp-server-ppt/issues/201)
- FC18: [#202 - Master-shape text access](https://github.com/trsdn/mcp-server-ppt/issues/202)
- FC19: [#203 - Picture and inherited backgrounds](https://github.com/trsdn/mcp-server-ppt/issues/203)
- FC20: [#204 - Table-cell text formatting](https://github.com/trsdn/mcp-server-ppt/issues/204)
- FC21: [#205 - Chart-type changes](https://github.com/trsdn/mcp-server-ppt/issues/205)
- FC22: [#206 - Picture transparent-color control](https://github.com/trsdn/mcp-server-ppt/issues/206)
- FC23: [#207 - SmartArt layout, style, and hierarchy changes](https://github.com/trsdn/mcp-server-ppt/issues/207)
- FC24: [#208 - Animation inspection and editing](https://github.com/trsdn/mcp-server-ppt/issues/208)
- FC25: [#209 - Media playback settings](https://github.com/trsdn/mcp-server-ppt/issues/209)
- FC26: [#210 - Slide/shape inspection](https://github.com/trsdn/mcp-server-ppt/issues/210)
- FC27: [#211 - Text diagnostics](https://github.com/trsdn/mcp-server-ppt/issues/211)
- FC28: [#212 - Hyperlink inventory and validation](https://github.com/trsdn/mcp-server-ppt/issues/212)
- FC29: [#213 - Proofing controls](https://github.com/trsdn/mcp-server-ppt/issues/213)
- FC30: [#214 - Slide/deck-wide find and replace](https://github.com/trsdn/mcp-server-ppt/issues/214)
- FC31: [#215 - Video export](https://github.com/trsdn/mcp-server-ppt/issues/215)
- FC32: [#216 - Text and image extraction](https://github.com/trsdn/mcp-server-ppt/issues/216)
- FC33: [#217 - Table/row convenience operations](https://github.com/trsdn/mcp-server-ppt/issues/217)
- FC34: [#218 - Notes convenience operations](https://github.com/trsdn/mcp-server-ppt/issues/218)
- FC35: [#219 - Transition convenience operations](https://github.com/trsdn/mcp-server-ppt/issues/219)
- FC36: [#220 - Cross-session slide copying](https://github.com/trsdn/mcp-server-ppt/issues/220)

## Later Options, Not Archive Blockers

### Theme Font Inspection

Potential follow-up to PR 5: read actual theme major/minor font roles, not just effective title/body placeholder formatting, which the target already supports. Define script coverage and unresolved font entries first. The local `GetThemeFonts` is a starting reference, not verified target-ready code; its chained font item access also needs explicit ownership review. Require independent COM read-back and a multi-design fixture. Keep font mutation out.

### Run-Level Formatting and Hyperlinks

This is explicitly listed as later work in Stefan's [native-feature audit](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/docs/POWERPOINT-NATIVE-FEATURE-AUDIT.md). It is not an existing local capability ready to port: local formatting also operates on whole shapes. Offer a dedicated design discussion around `TextRange.Characters`, 1-based ranges, mixed-format getters, bounds, Unicode, and preserving formatting outside the selection. Do not casually add range parameters to every existing action in one PR.

### Small Visual Evaluation Harness

Offer only after the recipes are accepted: four synthetic tasks, frozen skill/model inputs, exported PNGs, run metadata, and a structured outcome report. Reuse selected ideas from [eval/README.md](../eval/README.md), not the entire framework or official agent client. Confirm Stefan's preferred runner first; targeted code search did not find an existing judge/archetype harness, but that is not proof that none exists outside the inspected surfaces.

Keep builder and judge contexts separate, record failures separately from quality scores, and make cost/time budgets explicit. An LLM score must not stand in for a deterministic assertion that the correct text/data reached the file. Do not claim the recipes improve output until a fixed comparison demonstrates it. Automatic skill rewriting, private corpora, and making paid model evaluation mandatory in CI are out of scope.

### Other Capability Candidates

The [Feature Contribution Checklist](#feature-contribution-checklist) supersedes this earlier unreviewed shortlist with a source-level comparison. Confirm the current target gap, workflow value, and implementation behavior before promoting any entry to an implementation issue; the comparison is not runtime verification or maintainer acceptance.

## Contributions to Avoid

| Candidate | Decision and reason |
| --- | --- |
| Whole local COM/session layer | Do not port. Typed PIA conventions, shared fixtures, and shutdown policy differ materially. |
| Local fast teardown as a performance PR | Do not offer the process-kill design. Measurements here do not establish a compatible improvement there. |
| Full agent client and auto-tuning pipeline | Not in the initial proposal. Adds another runtime, model access, dependencies, and substantial maintenance ownership. |
| Entire learned-reference catalog | Do not transfer without rights/provenance review. It is unnecessary for four curated recipes. |
| CLI/MCP generation or skill synchronization | Already implemented in the target. Some target instructions are stale; that is a documentation correction, not a missing architecture. |
| Media insertion, tags, Save As, chart quick styles, linked pictures | Already delivered in the inspected target. Do not offer duplicates under different names. |
| Local document-property ordinal fixes | The target exposes named properties; the local wrong-ordinal bug does not establish a target bug. |
| Local BGR fixes as a blanket target bug report | Target text color output explicitly documents native `0xBBGGRR`. A different documented representation is not itself a reversal bug. |
| Physical printing, window controls, open-file passwords | Contrary to the target's explicitly documented scope/safety decisions. |
| Hook installer copied verbatim | Target hook has a `pwsh` shebang, unlike the original local failure case. Reproduce an actual Git-for-Windows failure before proposing a fix. Do not switch the target to Windows PowerShell 5 merely because the local installer does. |
| Local release/registry/packaging patches | Verify each against the target's release metadata and channels first; similarly named infrastructure is not evidence of the same bug. |
| Mandatory local full-suite pre-push evidence | A workflow-policy decision for Stefan, not a prerequisite for accepting a correctness PR. |

## Maintainer Message

The following message is suitable as an initial coordination note, not a claim that the implementation is ready:

> Hi Stefan, I'd like to retire my PowerPoint repo and contribute selected work to yours rather than maintain a parallel implementation. I compared the current surfaces and would keep your typed-PIA architecture, public contracts, and graceful shutdown policy.
>
> I suggest starting with the Core test-discovery fix in #71 and a separate improvement to the COM leak audit. The audit currently accepts a file with an unreleased acquisition if an unrelated release appears elsewhere in the same file; I verified that predicate with a small static probe. I'd also like to help investigate #72 at the partial-startup ownership boundary, without adding force-kill to normal shutdown.
>
> For a second batch, I can contribute theme-palette inspection and four focused composition recipes adapted to your existing skills. These would not require a new agent runtime or my full design catalog. Bounded text find/replace would be a later, separately agreed feature.
>
> Does that order fit your priorities, and is anyone already working on #71 or #72? I can keep each PR independently reviewable, with its own reproduction, focused tests, and the relevant CLI/MCP and documentation updates.

## Submission and Retirement Process

**Mandatory before any PR submission:** Implement and validate in a separate checkout based on Stefan's repository, using a feature branch in that clone or a fork of his repository. Build and run the relevant tests there against the exact changes to be submitted. Tests passing in this old repository are not evidence that the contribution works in his. Do not open a PR with untested changes, zero executed tests, or unresolved relevant failures; report the blocker instead. See Implementation Evidence above for the current target checkout and validation status.

1. Agree the first batch and check current issue/PR ownership. At inspection time #71 and #72 were open and no PRs were open; this can change immediately.
2. Work from a fresh fork/clone of Stefan's PowerPoint repository. Keep this repository's remote intact. Do not merge unrelated histories or repoint the current checkout as a shortcut.
3. In that target checkout, read its current contribution instructions, establish a focused baseline, and implement one outcome per branch. Reproduce the reported defect before fixing it; for new features agree the contract and write real-PowerPoint behavioral tests first. Build and rerun the relevant tests after implementation before submitting the PR.
4. In each PR explain the user/developer problem, exact root cause or API gap, chosen approach, rejected alternatives, compatibility, risks, and tests actually run. Link the local source/commit where substantial material was adapted and preserve its attribution.
5. Report the target base revision and tested contribution revision, build status, selected test command/filter, executed count, failures/skips, Office/.NET versions, and process-cleanup observations where relevant. An empty run is not passing evidence. Do not populate a PR with proposed tests presented as completed tests. Any subsequent code change requires rerunning the affected checks before the updated PR is represented as verified.
6. Track each candidate as accepted, deferred, or declined. A declined optional capability is a valid outcome, not a reason to keep both products alive indefinitely.
7. Before archiving, publish a short migration notice linking to Stefan's project and its installation instructions. State that commands, tool names, identifiers, package names, and session behavior are not drop-in compatible. Do not claim the target is a compatible update of existing packages.
8. Check the two CLI installations for command-name collisions: both document `pptcli`. Remove or explicitly disambiguate the old installation and MCP configuration as part of migration. Stop only proven-owned old daemon sessions, preserving user presentations.
9. Preserve commit history, tags, releases, applicable attribution, and any desired private assets in their proper private location. Do not copy private eval inputs into public retirement notes or PRs.
10. Resolve or label remaining local issues with an honest disposition and link only to agreed target work. Do not bulk-create duplicate target issues. Update local distribution notices only for channels that actually exist.
11. Disable obsolete publishing schedules/automation and review repository-scoped credentials before the archive step, with explicit authorization. Do not remove user-level credentials or unrelated integrations.
12. Archive only on Torsten's explicit instruction after the migration notice and distribution decisions are complete. Full feature parity and completion of the optional PR list are not required. Contributions can continue from the new fork after this repository is archived.

## Remaining Decisions and Risks

- **Stefan, before starting PRs 1/3:** confirm nobody already owns the two open issues. Welcome for contributions is not approval of every proposed implementation.
- **Stefan and Torsten, before PR 4:** agree structured test evidence and timeout defaults. This must not become an unsolicited testing-policy overhaul.
- **Stefan, before PRs 5/7:** approve selector/result contracts and replacement formatting semantics. These proposals are not zero-question implementation specs yet.
- **Torsten, before PR 6 or any asset transfer:** confirm reuse rights and exclude private learned examples. Public code licensing does not prove every source deck was redistributable.
- **Torsten, before archive:** decide which installed/published channels need retirement notices and authorize the archive operation separately.

The main uncertainties are target runtime behavior on this host, acceptance of the feature contracts, and the actual benefit of richer design guidance. The immediate defects have stronger evidence than the optional features. Re-rank if the target fixes #71/#72, changes its lifetime policy, adds palette/search APIs, or adopts a different evaluation framework before work begins.