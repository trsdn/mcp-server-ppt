# Fails when llm-tests reference an MCP tool that does not exist.
#
# Why this gate exists (#136): the llm-tests were inherited from an upstream Excel
# project and kept naming its tools - range, range_format, table, chart_config,
# screenshot. None exist on this server. The damage is not a broken test, it is a
# silently vacuous one:
#
#   allowed_tools=["chart", "table", "file", "range", "slide"]
#
# restricted the model to a set where three of five were fictional, and
#
#   assert result.tool_was_called("table")
#
# could never be true. Nobody noticed, because these tests are gated behind
# inputs.run_llm_gate inside a job gated on an unset variable targeting a runner
# that does not exist (#143) - so they have never executed.
#
# A rename fixes today. This gate fixes tomorrow.
#
# Ground truth is the GENERATED tool surface, the same source check-documented-counts.ps1
# uses, so this cannot drift from what the server actually exposes.
#
# ASCII only - see scripts/check-test-filters.ps1 for why.

[CmdletBinding()]
param(
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

$rootDir  = Split-Path $PSScriptRoot -Parent
$coreObj  = Join-Path $rootDir 'src\PptMcp.Core\obj'
$mcpTools = Join-Path $rootDir 'src\PptMcp.McpServer\Tools'
$llmTests = Join-Path $rootDir 'llm-tests'

if (-not $Quiet) {
    Write-Host 'LLM Test Tool Name Verification'
    Write-Host '==============================='
    Write-Host ''
}

foreach ($required in @($coreObj, $mcpTools, $llmTests)) {
    if (-not (Test-Path $required)) {
        Write-Host "ERROR: expected path not found: $required" -ForegroundColor Red
        Write-Host '       Build the solution first - the generated tool surface is the source of truth.' -ForegroundColor Red
        exit 1
    }
}

# --- Ground truth: the generated tool surface -----------------------------------

$known = [System.Collections.Generic.HashSet[string]]::new()

Get-ChildItem -Path $coreObj -Recurse -Filter 'ServiceRegistry.*.g.cs' |
    Group-Object Name | ForEach-Object { $_.Group[0] } | ForEach-Object {
        if ($_.Name -match '^ServiceRegistry\.(?<tool>.+)\.g\.cs$') {
            $candidate = $Matches['tool'].ToLowerInvariant()
            # The generator also emits ServiceRegistry.<tool>.dispatch.g.cs plus
            # DispatchHelpers/DispatchTable. Only the dotless names are tools; without
            # this filter the allow-list doubles and the gate stops discriminating.
            if ($candidate -notmatch '\.' -and $candidate -notmatch '^dispatch') {
                [void]$known.Add($candidate)
            }
        }
    }

foreach ($file in Get-ChildItem -Path $mcpTools -Filter '*.cs') {
    $content = Get-Content $file.FullName -Raw
    if ($content -notmatch '\[McpServerToolType\]') { continue }
    if ($content -match '\[McpServerTool\(Name\s*=\s*"(?<name>[^"]+)"') {
        [void]$known.Add($Matches['name'].ToLowerInvariant())
    }
}

if ($known.Count -eq 0) {
    Write-Host 'FAIL: found no tools in the generated surface.' -ForegroundColor Red
    Write-Host '      Refusing to report success without evidence. Build the solution first.'
    exit 1
}

# --- What the llm-tests actually name -------------------------------------------
#
# Scoped deliberately to allowed_tools=[...] and tool_was_called("..."), because a
# blanket scan would flag ordinary Python - `for i in range(20)` in cli/conftest.py
# is not a tool reference.

$referenced = @{}   # tool name -> list of "file:line"
$sites = 0

$pyFiles = Get-ChildItem -Path $llmTests -Recurse -Filter '*.py' |
    Where-Object { $_.FullName -notmatch '\\\.venv\\|\\__pycache__\\' }

foreach ($file in $pyFiles) {
    $lines = Get-Content $file.FullName
    $rel = $file.FullName.Substring($rootDir.Length + 1)

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $names = @()

        if ($line -match 'allowed_tools\s*=\s*\[(?<inner>[^\]]*)\]') {
            $names += [regex]::Matches($Matches['inner'], '"(?<n>[^"]+)"') |
                ForEach-Object { $_.Groups['n'].Value }
        }
        $names += [regex]::Matches($line, 'tool_was_called\("(?<n>[^"]+)"\)') |
            ForEach-Object { $_.Groups['n'].Value }

        foreach ($n in $names) {
            $sites++
            $key = $n.ToLowerInvariant()
            if (-not $referenced.ContainsKey($key)) { $referenced[$key] = @() }
            $referenced[$key] += "$rel`:$($i + 1)"
        }
    }
}

# --- Refuse to pass without evidence --------------------------------------------

if ($sites -eq 0) {
    Write-Host 'FAIL: inspected the llm-tests and found no tool references at all.' -ForegroundColor Red
    Write-Host '      Either the tests moved or their shape changed. A gate that finds'
    Write-Host '      nothing must fail, or it silently stops guarding anything.'
    exit 1
}

$unknown = @($referenced.Keys | Where-Object { -not $known.Contains($_) } | Sort-Object)

if (-not $Quiet) {
    Write-Host "Tools on the server:      $($known.Count)"
    Write-Host "Tool references in tests: $sites across $(@($referenced.Keys).Count) distinct name(s)"
    Write-Host ''
}

if ($unknown.Count -gt 0) {
    Write-Host "FAIL: $($unknown.Count) tool name(s) referenced by llm-tests do not exist." -ForegroundColor Red
    Write-Host ''
    foreach ($u in $unknown) {
        Write-Host "  $u" -ForegroundColor Red
        foreach ($site in $referenced[$u]) { Write-Host "      $site" }
    }
    Write-Host ''
    Write-Host 'A test that allows only nonexistent tools, or asserts a nonexistent tool was'
    Write-Host 'called, cannot pass. Use a name from the generated surface:'
    Write-Host ''
    Write-Host "  $((@($known) | Sort-Object) -join ', ')"
    exit 1
}

if (-not $Quiet) {
    Write-Host "OK: all $(@($referenced.Keys).Count) referenced tool name(s) exist on the server." -ForegroundColor Green
}
exit 0
