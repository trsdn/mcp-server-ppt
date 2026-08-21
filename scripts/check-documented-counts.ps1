# Verifies that the tool and operation counts stated in the documentation match
# the tool surface the source generators actually produce.
#
# This gate exists because every documented count had silently drifted: FEATURES.md
# claimed 204 operations while its own table summed to 104, and the extension
# manifest claimed 25 tools with 225 operations. Nothing verified any of it.
#
# The gate fails when it detects nothing. A coverage check that reports success
# without inspecting anything is worse than no check at all.

$ErrorActionPreference = 'Stop'

$rootDir = Split-Path -Parent $PSScriptRoot
# The generated service registry in the Core project is the ground truth for the
# tool surface: one registry entry per tool, one enum member per action. Only the
# Core project needs to have been built, so this runs in any build job.
$coreObj = Join-Path $rootDir 'src\PptMcp.Core\obj'
$mcpTools = Join-Path $rootDir 'src\PptMcp.McpServer\Tools'

Write-Host 'Documented Count Verification'
Write-Host '============================='
Write-Host ''

foreach ($required in @($coreObj, $mcpTools)) {
    if (-not (Test-Path $required)) {
        Write-Host "ERROR: expected path not found: $required" -ForegroundColor Red
        Write-Host '       Build the solution first - the generated tool surface is the source of truth.' -ForegroundColor Red
        exit 1
    }
}

# --- Ground truth: the generated service registry -------------------------------

$registryFiles = Get-ChildItem -Path $coreObj -Recurse -Filter 'ServiceRegistry.*.g.cs' |
    Group-Object Name | ForEach-Object { $_.Group[0] }

$actionsByTool = @{}
foreach ($file in $registryFiles) {
    if ($file.Name -notmatch '^ServiceRegistry\.(?<tool>.+)\.g\.cs$') { continue }
    $tool = $Matches['tool'].ToLowerInvariant()
    $content = Get-Content $file.FullName -Raw
    $actions = [regex]::Matches($content, 'JsonStringEnumMemberName\("(?<a>[^"]+)"\)') |
        ForEach-Object { $_.Groups['a'].Value } |
        Sort-Object -Unique
    if ($actions.Count -gt 0) { $actionsByTool[$tool] = @($actions) }
}

# --- The MCP tool surface -------------------------------------------------------
#
# Hand-written tools declare their actions inline rather than through the registry,
# so their operations are counted from the tool's own action enum.
$handWritten = @{}
foreach ($file in Get-ChildItem -Path $mcpTools -Filter '*.cs') {
    $content = Get-Content $file.FullName -Raw
    if ($content -notmatch '\[McpServerToolType\]') { continue }
    if ($content -notmatch '\[McpServerTool\(Name\s*=\s*"(?<name>[^"]+)"') { continue }
    $name = $Matches['name'].ToLowerInvariant()
    $actions = [regex]::Matches($content, 'JsonStringEnumMemberName\("(?<a>[^"]+)"\)') |
        ForEach-Object { $_.Groups['a'].Value } |
        Sort-Object -Unique
    if ($actions.Count -gt 0) { $handWritten[$name] = @($actions) }
}

$surface = @{}
foreach ($tool in $actionsByTool.Keys) { $surface[$tool] = $actionsByTool[$tool] }
# A hand-written MCP tool replaces the generated registry entry of the same name,
# because it is the hand-written tool that defines the exposed action set.
foreach ($tool in $handWritten.Keys) { $surface[$tool] = $handWritten[$tool] }

$toolCount = $surface.Count
$opCount = ($surface.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum

# --- Refuse to pass without evidence -------------------------------------------

if ($toolCount -eq 0 -or $opCount -eq 0) {
    Write-Host "ERROR: detected $toolCount tools and $opCount operations." -ForegroundColor Red
    Write-Host '       Detecting nothing means this check inspected nothing, so it cannot pass.' -ForegroundColor Red
    Write-Host '       Build the solution, then re-run.' -ForegroundColor Red
    exit 1
}

Write-Host "Detected tool surface: $toolCount tools, $opCount operations" -ForegroundColor Green
Write-Host ''

# --- Compare against every documented count ------------------------------------

# --- Compare against every documented count ------------------------------------
#
# Scanned repository-wide rather than through a fixed file list, so a newly added
# document that states a count is covered automatically.

$skipDirs = '^(node_modules|\.git|obj|bin)\\|\\(node_modules|\.git|obj|bin)\\'
$docs = Get-ChildItem -Path $rootDir -Recurse -Filter '*.md' -File |
    Where-Object { $_.FullName.Substring($rootDir.Length + 1) -notmatch $skipDirs } |
    # Changelogs deliberately record superseded counts, so they are not current claims.
    Where-Object { $_.Name -ne 'CHANGELOG.md' }

# Each pattern captures a tool count, an operation count, or both.
$patterns = @(
    '(?<t>\d+)\s+(?:specialized\s+)?(?:MCP\s+)?tools?\s+with\s+(?<o>\d+)\s+operations'
    '(?<t>\d+)\s+command categories with\s+(?<o>\d+)\s+operations'
    '\*\*(?<o>\d+) operations\*\* across (?<t>\d+) command categories'
    '\*\*Total \((?<t>\d+) tools\)\*\* \| \*\*(?<o>\d+)\*\*'
    'all (?<o>\d+) operations'
)

$failures = @()
$checked = 0

foreach ($doc in $docs) {
    $content = Get-Content $doc.FullName -Raw
    $relative = $doc.FullName.Substring($rootDir.Length + 1)
    foreach ($pattern in $patterns) {
        foreach ($m in [regex]::Matches($content, $pattern)) {
            $checked++
            if ($m.Groups['t'].Success -and [int]$m.Groups['t'].Value -ne $toolCount) {
                $failures += "{0}: states {1} tools, actual is {2}  (in '{3}')" -f `
                    $relative, $m.Groups['t'].Value, $toolCount, $m.Value.Trim()
            }
            if ($m.Groups['o'].Success -and [int]$m.Groups['o'].Value -ne $opCount) {
                $failures += "{0}: states {1} operations, actual is {2}  (in '{3}')" -f `
                    $relative, $m.Groups['o'].Value, $opCount, $m.Value.Trim()
            }
        }
    }
}

# FEATURES.md states a total and also lists a count per tool. The header once said
# 204 while the table summed to 104, so the two are reconciled explicitly.
$featuresPath = Join-Path $rootDir 'FEATURES.md'
if (Test-Path $featuresPath) {
    $featuresContent = Get-Content $featuresPath -Raw
    $rows = [regex]::Matches($featuresContent, '(?m)^\|\s*(?:\*\*)?`?(?<tool>[a-z][a-z0-9]*)`?(?:\*\*)?\s*\|\s*(?<n>\d+)\s*\|')
    if ($rows.Count -eq 0) {
        $failures += 'FEATURES.md: per-tool operation table not found, so its total cannot be reconciled'
    }
    else {
        $checked++
        $sum = ($rows | ForEach-Object { [int]$_.Groups['n'].Value } | Measure-Object -Sum).Sum
        if ($sum -ne $opCount) {
            $failures += "FEATURES.md: per-tool table sums to $sum, actual is $opCount"
        }
        if ($rows.Count -ne $toolCount) {
            $failures += "FEATURES.md: per-tool table lists $($rows.Count) tools, actual is $toolCount"
        }
    }
}

if ($checked -eq 0) {
    Write-Host 'ERROR: no documented counts were found to verify.' -ForegroundColor Red
    Write-Host '       Either the documentation stopped stating counts or the patterns are stale.' -ForegroundColor Red
    exit 1
}

if ($failures.Count -gt 0) {
    Write-Host "Documented counts do not match the generated tool surface:" -ForegroundColor Red
    Write-Host ''
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host ''
    Write-Host "Expected: $toolCount tools, $opCount operations" -ForegroundColor Yellow
    exit 1
}

Write-Host "All $checked documented counts match the generated tool surface." -ForegroundColor Green
exit 0
