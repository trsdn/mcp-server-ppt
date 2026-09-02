# Verifies that the hand-written session/file management surface exposes the same
# action set on both entry points.
#
# Every other tool is generated from a Core interface, so CLI/MCP parity is
# structural. Session management is the one exception (PptFileTool.cs says so in a
# comment), and it is the bootstrap layer every workflow calls first - so the one
# surface with no construction-time guarantee is the one that cannot afford drift.
# It had already diverged: MCP exposed 'test', the CLI did not (issue #131).
#
# check-documented-counts.ps1 does not cover this axis. It counts the MCP surface,
# so a CLI that is missing an action is invisible to it.
#
# ASCII only. Windows PowerShell reads these files as cp1252, and a single non-ASCII
# byte can turn the whole script into a parse error, silently disabling the gate.

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$mcpFile = Join-Path $repoRoot 'src\PptMcp.McpServer\Tools\PptFileTool.cs'
$cliFile = Join-Path $repoRoot 'src\PptMcp.CLI\Program.cs'

Write-Host 'Session/file parity check'
Write-Host '========================='
Write-Host ''

foreach ($f in @($mcpFile, $cliFile)) {
    if (-not (Test-Path $f)) {
        Write-Host "FAIL: expected source file not found: $f" -ForegroundColor Red
        exit 1
    }
}

# MCP: the JsonStringEnumMemberName on each PptFileAction member is the wire action.
$mcpText = Get-Content $mcpFile -Raw
$enumBlock = [regex]::Match($mcpText, 'enum\s+PptFileAction\s*\{(?<body>[^}]*)\}')

if (-not $enumBlock.Success) {
    Write-Host 'FAIL: could not locate the PptFileAction enum in PptFileTool.cs.' -ForegroundColor Red
    Write-Host '      The gate cannot verify parity it cannot read.' -ForegroundColor Red
    exit 1
}

$mcpActions = [regex]::Matches($enumBlock.Groups['body'].Value, 'JsonStringEnumMemberName\("(?<name>[^"]+)"\)') |
    ForEach-Object { $_.Groups['name'].Value } |
    Sort-Object -Unique

# CLI: the string literal passed to AddCommand inside the "session" branch.
$cliText = Get-Content $cliFile -Raw
$branch = [regex]::Match($cliText, 'AddBranch\("session",\s*branch\s*=>\s*\{(?<body>.*?)\r?\n\s*\}\);', 'Singleline')

if (-not $branch.Success) {
    Write-Host 'FAIL: could not locate the "session" branch registration in Program.cs.' -ForegroundColor Red
    Write-Host '      The gate cannot verify parity it cannot read.' -ForegroundColor Red
    exit 1
}

$cliActions = [regex]::Matches($branch.Groups['body'].Value, 'AddCommand<[^>]+>\("(?<name>[^"]+)"\)') |
    ForEach-Object { $_.Groups['name'].Value } |
    Sort-Object -Unique

# A gate that inspects nothing must fail rather than report success without evidence.
if ($mcpActions.Count -eq 0) {
    Write-Host 'FAIL: found 0 MCP file actions. The enum shape must have changed.' -ForegroundColor Red
    exit 1
}

if ($cliActions.Count -eq 0) {
    Write-Host 'FAIL: found 0 CLI session actions. The branch shape must have changed.' -ForegroundColor Red
    exit 1
}

Write-Host ("MCP file actions ({0}):     {1}" -f $mcpActions.Count, ($mcpActions -join ', '))
Write-Host ("CLI session actions ({0}):  {1}" -f $cliActions.Count, ($cliActions -join ', '))
Write-Host ''

$missingFromCli = @($mcpActions | Where-Object { $_ -notin $cliActions })
$missingFromMcp = @($cliActions | Where-Object { $_ -notin $mcpActions })

if ($missingFromCli.Count -eq 0 -and $missingFromMcp.Count -eq 0) {
    Write-Host ("Session/file parity holds: both entry points expose the same {0} actions." -f $mcpActions.Count) -ForegroundColor Green
    exit 0
}

Write-Host 'Session/file parity is broken:' -ForegroundColor Red

foreach ($a in $missingFromCli) {
    Write-Host ("  - '{0}' is in MCP 'file' but not in CLI 'session'" -f $a) -ForegroundColor Red
}

foreach ($a in $missingFromMcp) {
    Write-Host ("  - '{0}' is in CLI 'session' but not in MCP 'file'" -f $a) -ForegroundColor Red
}

Write-Host ''
Write-Host 'Both entry points are documented as first-class and equal. Add the missing'
Write-Host 'action to the other surface, or remove it from the one that has it.'
exit 1
