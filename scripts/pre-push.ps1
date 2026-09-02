# Pre-push hook: refuses to push code that no local integration run has covered.
#
# Install:
#   Copy-Item scripts\pre-push.ps1 .git\hooks\pre-push
#
# WHY A PRE-PUSH HOOK AND NOT PRE-COMMIT
#
# The integration suite takes tens of minutes. Running it per commit would make people
# stop committing. Running it per push matches what a push means: "this is ready to be
# seen". Evidence is bound to a SHA, so one gate run covers any number of pushes of that
# same commit.
#
# WHY IT IS BLOCKING
#
# The powerpoint-integration CI job cannot run and will not be provisioned, so no
# automated check downstream of this point ever executes a test against real PowerPoint.
# If this hook only warned, nothing in the repository would enforce integration coverage
# at all - which is the state issue #143 describes.
#
# THE OVERRIDE IS DELIBERATELY LOUD
#
#   $env:PPTMCP_SKIP_INTEGRATION_GATE = '1'
#
# It prints a banner naming the commit it waved through, so an override is visible in the
# push output and in any terminal transcript rather than being a silent habit.
#
# ASCII only - see scripts/check-test-filters.ps1 for why.

$ErrorActionPreference = 'Stop'

# git passes "<localRef> <localSha> <remoteRef> <remoteSha>" lines on stdin.
$refLines = @($input | Where-Object { $_ -and $_.Trim() -ne '' })

$zero = '0' * 40

# Nothing on stdin means nothing to push.
if ($refLines.Count -eq 0) { exit 0 }

$repoRoot = (git rev-parse --show-toplevel).Trim()
Set-Location $repoRoot

$codePaths = @('src/', 'tests/')
$needsEvidence = @()

foreach ($line in $refLines) {
    $parts = $line.Trim() -split '\s+'
    if ($parts.Count -lt 4) { continue }

    $localSha  = $parts[1]
    $remoteSha = $parts[3]

    # Branch deletion.
    if ($localSha -eq $zero) { continue }

    if ($remoteSha -eq $zero) {
        # New branch on the remote: compare against the default branch so a long-lived
        # feature branch is judged by what it adds, not by the entire history.
        $base = (git merge-base $localSha origin/main 2>$null)
        if (-not $base) { $base = (git rev-list --max-parents=0 $localSha | Select-Object -Last 1) }
        $range = "$($base.Trim())..$localSha"
    }
    else {
        $range = "$remoteSha..$localSha"
    }

    $changed = @(git diff --name-only $range 2>$null)
    $touchesCode = @($changed | Where-Object {
        $path = $_
        ($codePaths | Where-Object { $path.StartsWith($_) }).Count -gt 0
    })

    if ($touchesCode.Count -gt 0) {
        $needsEvidence += [pscustomobject]@{ Sha = $localSha; Files = $touchesCode.Count }
    }
}

if ($needsEvidence.Count -eq 0) {
    # Docs, workflows, skills. Nothing that the integration suite would exercise.
    exit 0
}

if ($env:PPTMCP_SKIP_INTEGRATION_GATE -eq '1') {
    Write-Host ""
    Write-Host "###############################################################" -ForegroundColor Yellow
    Write-Host "#  INTEGRATION GATE OVERRIDDEN                                #" -ForegroundColor Yellow
    Write-Host "###############################################################" -ForegroundColor Yellow
    Write-Host ""
    foreach ($item in $needsEvidence) {
        Write-Host "  pushing $($item.Sha) - $($item.Files) code file(s) changed, UNVERIFIED against PowerPoint" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  PPTMCP_SKIP_INTEGRATION_GATE=1 was set. No integration test ran for" -ForegroundColor Yellow
    Write-Host "  this push, and no CI job will run one either - see issue #143." -ForegroundColor Yellow
    Write-Host "  Say so in the pull request." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

foreach ($item in $needsEvidence) {
    Write-Host ""
    Write-Host "Pre-push: checking integration evidence for $($item.Sha)" -ForegroundColor Cyan
    Write-Host "          ($($item.Files) file(s) under src/ or tests/ in this push)"

    & "$repoRoot\scripts\check-integration-evidence.ps1" -Commit $item.Sha
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "PUSH REFUSED." -ForegroundColor Red
        Write-Host ""
        Write-Host "  This push changes code that only a real PowerPoint can verify, and no"
        Write-Host "  CI job can run those tests. Produce evidence first:"
        Write-Host ""
        Write-Host "    .\scripts\Invoke-LocalIntegrationGate.ps1"
        Write-Host ""
        Write-Host "  To push anyway, deliberately and on the record:"
        Write-Host ""
        Write-Host "    `$env:PPTMCP_SKIP_INTEGRATION_GATE = '1'; git push; `$env:PPTMCP_SKIP_INTEGRATION_GATE = `$null"
        Write-Host ""
        exit 1
    }
}

Write-Host ""
Write-Host "Pre-push integration gate passed." -ForegroundColor Green
exit 0
