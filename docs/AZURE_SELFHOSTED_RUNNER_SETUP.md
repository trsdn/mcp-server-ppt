# Azure Self-Hosted Runner Setup

> **DECISION: this runner was considered and DECLINED.** It will not be provisioned.
> The document is kept as a record of that decision and of what activation would take,
> so the question is not silently re-litigated. See "Decision" below before planning any
> work that assumes CI runs the integration suite - it does not, and will not.

This document describes the minimum setup that *would* be needed to activate the
`powerpoint-integration` job in `.github/workflows/integration-tests.yml`.

## Decision: declined, with the coverage moved to a local gate

Real PowerPoint COM automation needs a licensed Office install and an interactive desktop
session. Providing that in CI means running, licensing, patching and securing a persistent
Windows host - a standing cost, and a machine holding repository credentials with a desktop
session permanently logged in. The maintainer decided that cost is not justified for this
project.

What replaced it: the suite runs on the maintainer's machine and is enforced **at push
time**, which is the only point where a human is present and PowerPoint is available.

```powershell
.\scripts\Invoke-LocalIntegrationGate.ps1   # runs the suite, writes an evidence manifest
.\scripts\install-hooks.ps1                 # installs the pre-push hook that requires it
```

The gate writes `.integration-evidence/manifest.json`, recording the commit SHA, every
suite, the number of tests that ran, and the result. `scripts/check-integration-evidence.ps1`
verifies it, and the pre-push hook refuses any push touching `src/` or `tests/` without
evidence **for that exact commit**. Evidence for a different SHA, or produced from a dirty
working tree, is rejected - stale evidence looks like coverage and is worse than none.

The override is `PPTMCP_SKIP_INTEGRATION_GATE=1`, and it prints a banner naming the
unverified commit so that skipping is visible rather than habitual.

The suite list in the local gate deliberately mirrors the steps of the unreachable
`powerpoint-integration` job, so the two cannot drift apart in what they claim to cover.

## What activation would take, if the decision is ever revisited

The workflow is still in the repository and becomes active when both of these are true:

- repository variable `ENABLE_POWERPOINT_INTEGRATION_CI` is set to `true`
- a self-hosted Windows runner with the label `powerpoint` is available

Until then, the workflow reports a status message instead of pretending that PowerPoint
integration is covered in CI.

## Recommended Host Requirements

- Windows 11 or Windows Server with desktop experience
- Microsoft 365 Apps / PowerPoint installed and licensed
- .NET SDK `9.0.x`
- `uv` available on PATH for `llm-tests/`
- Stable disk space for build outputs and test artifacts
- Runner labels: `self-hosted`, `windows`, `powerpoint`

## Desktop Session Requirement

PowerPoint COM automation is not reliably headless. Use a runner host that keeps an interactive desktop session available for the runner user.

Recommended practice:

- dedicate the machine to PowerPoint integration workloads
- use a dedicated local/service account for the runner
- verify that PowerPoint can open and close normally under that account before enabling CI

## Basic Setup Steps

1. Provision the Windows host or Azure VM.
2. Install PowerPoint and confirm it opens successfully for the runner account.
3. Install the .NET 9 SDK.
4. Install `uv`.
5. Register the GitHub Actions runner for this repository.
6. Add the `powerpoint` label to that runner.
7. Set repository variable `ENABLE_POWERPOINT_INTEGRATION_CI=true`.
8. Optionally add secret `AZURE_OPENAI_ENDPOINT` if you want workflow-dispatch LLM gate runs.

## Validation Checklist

Before enabling the repository variable, validate on the runner host:

```powershell
dotnet build src\PptMcp.CLI\PptMcp.CLI.csproj -c Release
dotnet build src\PptMcp.McpServer\PptMcp.McpServer.csproj -c Release
.\scripts\Test-CliWorkflow.ps1
dotnet test tests\PptMcp.McpServer.Tests\PptMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~McpServerIntegrationTests.SmokeTest_AllTools_E2EWorkflow"
```

If those pass locally on the runner host, enable `ENABLE_POWERPOINT_INTEGRATION_CI` and trigger `integration-tests.yml` with `workflow_dispatch`.

## Optional LLM Regression Gate

The workflow can also run the canonical LLM regression gate when dispatched manually.

Prerequisites:

- `AZURE_OPENAI_ENDPOINT` secret configured
- runner host already passes the regular PowerPoint smoke/integration steps

Manual local command:

```powershell
.\scripts\Test-LlmRegressionGate.ps1
```
