#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Tests the PowerPoint CLI end-to-end workflow - exactly what a user would do.

.DESCRIPTION
    This script demonstrates and tests a basic CLI workflow:
    1. Create session (auto-starts daemon, creates file)
    2. Create slide with Blank layout
    3. List slides
    4. Add textbox content
    5. List shapes on the slide
    6. Close session (with save)
    7. Reopen saved file (session open - exercises Presentations.Open path)
    8. List slides in reopened session
    9. List shapes in reopened session
    10. Close reopened session
    11. Verify file exists

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [switch]$KeepFile  # Don't delete the test file after completion
)

$ErrorActionPreference = 'Stop'

# Find CLI executable (prefer Release build)
$candidateCliPaths = @(
    "..\src\PptMcp.CLI\bin\Release\net9.0-windows\pptcli.exe",
    "..\src\PptMcp.CLI\bin\Debug\net9.0-windows\pptcli.exe",
    "..\src\PptMcp.CLI\bin\Release\net10.0-windows\pptcli.exe",
    "..\src\PptMcp.CLI\bin\Debug\net10.0-windows\pptcli.exe"
) | ForEach-Object { Join-Path $PSScriptRoot $_ }

$cliPath = $candidateCliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $cliPath) {
    Write-Error "CLI not found. Build first: dotnet build src/PptMcp.CLI"
    exit 1
}

$cli = (Resolve-Path $cliPath).Path
Write-Host "Using CLI: $cli" -ForegroundColor Cyan

# Generate unique test file
$testFile = Join-Path $env:TEMP "cli-workflow-test-$(Get-Random).pptx"
Write-Host "Test file: $testFile" -ForegroundColor Cyan

$passed = 0
$failed = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Action,
        [scriptblock]$Verify = $null
    )

    Write-Host "`n[$Name]" -ForegroundColor Yellow
    try {
        $result = & $Action
        if ($Verify) {
            $verifyResult = & $Verify $result
            if (-not $verifyResult) {
                Write-Host "  FAIL: Verification failed" -ForegroundColor Red
                Write-Host "  Result: $($result | ConvertTo-Json -Depth 5 -Compress)" -ForegroundColor Gray
                $script:failed++
                return $null
            }
        }
        Write-Host "  PASS" -ForegroundColor Green
        $script:passed++
        return $result
    }
    catch {
        Write-Host "  FAIL: $_" -ForegroundColor Red
        if ($_.ErrorDetails.Message) {
            Write-Host "  Details: $($_.ErrorDetails.Message)" -ForegroundColor Gray
        }
        $script:failed++
        return $null
    }
}

# ============================================================================
# TEST WORKFLOW
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PowerPoint CLI Workflow Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 1. Create session (auto-starts daemon, creates file)
$session = Test-Step "Create session (create file)" {
    & $cli -q session create $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -ne $false
}

if (-not $session.sessionId) {
    Write-Host "`nFATAL: Could not open session. Aborting." -ForegroundColor Red
    exit 1
}

$sessionId = $session.sessionId
Write-Host "  Session ID: $sessionId" -ForegroundColor Gray

# 2. Create slide
Test-Step "Create slide with Blank layout" {
    & $cli -q slide create --session $sessionId --position 0 --layout-name Blank | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 3. List slides
$slides = Test-Step "List slides" {
    & $cli -q slide list --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $null -ne $r.slides
}

Write-Host "  Slides: $(($slides.slides | Measure-Object).Count)" -ForegroundColor Gray

# 4. Add textbox content
Test-Step "Add textbox to slide 1" {
    & $cli -q shape add-textbox --session $sessionId --slide-index 1 --left 72 --top 72 --width 240 --height 48 --text "CLI smoke test" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 5. List shapes on slide 1
$shapes = Test-Step "List shapes on slide 1" {
    & $cli -q shape list --session $sessionId --slide-index 1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $null -ne $r.shapes
}

Write-Host "  Shapes: $(($shapes.shapes | Measure-Object).Count)" -ForegroundColor Gray

# 6. Close session (with save)
Test-Step "Close session (with save)" {
    & $cli -q session close --session $sessionId --save | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 7. Reopen saved file (session open - exercises Presentations.Open path distinct from Add+SaveAs)
#    This step would catch deployment issues like missing office.dll (issue #487) because
#    PptBatch.ctor runs AutomationSecurity setup before opening any presentation.
$reopenSession = Test-Step "Reopen saved file (session open)" {
    & $cli -q session open $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -ne $false
}

# 8. List slides in reopened session (proves the file loaded correctly)
if ($reopenSession -and $reopenSession.sessionId) {
    $reopenSessionId = $reopenSession.sessionId
    $reopenedSlides = Test-Step "List slides in reopened session" {
        & $cli -q slide list --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -and $null -ne $r.slides
    }

    Write-Host "  Reopened slides: $(($reopenedSlides.slides | Measure-Object).Count)" -ForegroundColor Gray

    # 9. List shapes in reopened session (proves saved content loaded correctly)
    $reopenedShapes = Test-Step "List shapes in reopened slide 1" {
        & $cli -q shape list --session $reopenSessionId --slide-index 1 | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -and $null -ne $r.shapes
    }

    Write-Host "  Reopened shapes: $(($reopenedShapes.shapes | Measure-Object).Count)" -ForegroundColor Gray

    # 10. Close reopened session
    Test-Step "Close reopened session" {
        & $cli -q session close --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true
    }
}

# 11. Verify file exists
Test-Step "Verify file exists" {
    if (Test-Path $testFile) {
        $size = (Get-Item $testFile).Length
        "File size: $size bytes"
    } else {
        throw "File not found"
    }
} -Verify {
    param($r)
    $r -match "bytes"
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $passed" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "Test file: $testFile" -ForegroundColor Gray

if (-not $KeepFile -and (Test-Path $testFile)) {
    Remove-Item $testFile -Force
    Write-Host "(Test file deleted)" -ForegroundColor Gray
} elseif ($KeepFile) {
    Write-Host "(Test file kept for inspection)" -ForegroundColor Yellow
}

if ($failed -gt 0) {
    Write-Host "`nSome tests FAILED!" -ForegroundColor Red
    exit 1
} else {
    Write-Host "`nAll tests PASSED!" -ForegroundColor Green
    exit 0
}
