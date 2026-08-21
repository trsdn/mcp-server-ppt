#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks that all CLI Settings properties are used in the args switch statements.

.DESCRIPTION
    Detects the pattern where a developer adds a Setting property but forgets to include it
    in the args switch statement, causing user values to be silently dropped.

    Example bug detected:
    - Settings has: public string? ConnectionString { get; init; }
    - Switch case has: new { connectionName, refreshOnOpen } // connectionString missing!

.EXAMPLE
    .\check-cli-settings-usage.ps1

.NOTES
    Part of pre-commit checks. Fails if any CLI command has unused Settings properties.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$cliCommandsDir = Join-Path $rootDir "src\PptMcp.CLI\Commands"

# Properties that are legitimately not passed to daemon (session management, meta properties)
$globalExclusions = @(
    "Action",
    "SessionId"
)

# Properties that are defined for future features but daemon doesn't support yet
# These should be reviewed periodically to see if they can be implemented
$futureFeatureExclusions = @(
    "SheetScope",      # NamedRange - daemon doesn't support worksheet-scoped names yet
    "ModuleType",      # VBA - daemon uses auto-detection for module type
    "LayoutStyle",     # PivotTable - uses LayoutType instead
    "TargetPivotTableName", # Slicer - not implemented in daemon
    "Position",        # Chart - parsed from position string, not passed directly
    "TargetSheet",     # Chart - uses SheetName instead for placement
    "EnableRefresh"    # Connection - daemon uses RefreshOnFileOpen instead
)

# Properties that are used indirectly (files read into other properties)
$indirectUsagePatterns = @{
    "ConnectionStringFile" = "ConnectionString"
    "CommandTextFile" = "CommandText"
    "MCodeFile" = "MCode"
    "CodeFile" = "Code"
    "CsvFile" = "CsvData"
    "DaxQueryFile" = "DaxQuery"
    "DmvQueryFile" = "DmvQuery"
    "ExpressionFile" = "Expression"
    "ValuesFile" = "Values"
    "FormulasFile" = "Formulas"
    "FormatsFile" = "Formats"
}

$issues = @()
$totalChecked = 0
$totalPassed = 0

function Get-SettingsProperties {
    param([string]$content)

    $properties = @()
    # The Settings class body must be delimited by brace matching. A regex that
    # captures to end of file also swallows every type declared after Settings,
    # which made JSON DTOs such as BatchEntry and BatchResult look like unused
    # settings properties.
    #
    # A file may declare several Settings classes (SessionCommands.cs declares
    # four), so every occurrence is read, not just the first.
    foreach ($header in [regex]::Matches($content, 'internal sealed class Settings[^{]*\{')) {
        $start = $header.Index + $header.Length
        $depth = 1
        $i = $start
        while ($i -lt $content.Length -and $depth -gt 0) {
            $ch = $content[$i]
            if ($ch -eq '{') { $depth++ }
            elseif ($ch -eq '}') { $depth-- }
            $i++
        }
        if ($depth -ne 0) {
            throw "Unbalanced braces while reading a Settings class body."
        }

        $settingsBlock = $content.Substring($start, $i - $start - 1)
        foreach ($match in [regex]::Matches($settingsBlock, 'public\s+\w+\??\s+(\w+)\s*\{')) {
            $properties += $match.Groups[1].Value
        }
    }
    return $properties
}

function Get-UsedProperties {
    param([string]$content)

    $usedProps = @()
    # Find all settings.PropertyName usages
    $usageMatches = [regex]::Matches($content, 'settings\.(\w+)')
    foreach ($match in $usageMatches) {
        $usedProps += $match.Groups[1].Value
    }
    return $usedProps | Sort-Object -Unique
}

Write-Host "Checking CLI Settings property usage..." -ForegroundColor Cyan
Write-Host ""

# Hand-written command files are named both *Command.cs and *Commands.cs, so the
# previous "*Command.cs" filter silently excluded DiagCommands, ServiceCommands
# and SessionCommands from the check.
$commandFiles = Get-ChildItem -Path $cliCommandsDir -Filter "*.cs" -File

foreach ($file in $commandFiles) {
    # Skip ListActionsCommand - it's a meta command
    if ($file.Name -eq "ListActionsCommand.cs") {
        continue
    }

    $content = Get-Content $file.FullName -Raw
    $fileName = $file.Name

    # Skip if no Settings class
    if (-not ($content -match 'internal sealed class Settings')) {
        continue
    }

    $totalChecked++

    $settingsProps = Get-SettingsProperties $content
    $usedProps = Get-UsedProperties $content

    $unusedProps = @()
    foreach ($prop in $settingsProps) {
        # Skip global exclusions
        if ($globalExclusions -contains $prop) {
            continue
        }

        # Skip future feature exclusions
        if ($futureFeatureExclusions -contains $prop) {
            continue
        }

        # Skip indirect usage (file properties that populate other properties)
        if ($indirectUsagePatterns.ContainsKey($prop)) {
            continue
        }

        # Check if property is used
        if ($usedProps -notcontains $prop) {
            $unusedProps += $prop
        }
    }

    if ($unusedProps.Count -gt 0) {
        $issues += [PSCustomObject]@{
            File = $fileName
            UnusedProperties = $unusedProps -join ", "
        }
    }
    else {
        $totalPassed++
    }
}

if ($totalChecked -eq 0) {
    Write-Host "ERROR: no CLI command files with a Settings class were inspected." -ForegroundColor Red
    Write-Host "       Detecting nothing means this check inspected nothing, so it cannot pass." -ForegroundColor Red
    exit 1
}

if ($issues.Count -gt 0) {
    Write-Host "Found CLI commands with unused Settings properties:" -ForegroundColor Red
    Write-Host ""

    foreach ($issue in $issues) {
        Write-Host "   $($issue.File)" -ForegroundColor Yellow
        Write-Host "      Unused: $($issue.UnusedProperties)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "   These Settings properties are defined but NOT passed to daemon in args." -ForegroundColor Red
    Write-Host "   User values will be silently ignored!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Fix: Add property to args switch statement:" -ForegroundColor Cyan
    Write-Host "        ""action"" => new { ..., propertyName = settings.PropertyName }," -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "CLI Settings usage check passed - $totalPassed/$totalChecked commands" -ForegroundColor Green
Write-Host "   All Settings properties are used in args switch statements" -ForegroundColor Gray
exit 0
