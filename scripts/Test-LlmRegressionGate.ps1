#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs the canonical manual LLM regression gate.

.DESCRIPTION
    Builds the CLI and MCP Server in Release mode, points the LLM tests at those
    binaries, and runs the standard three CLI + three MCP pytest scenarios.

    This gate stays manual/on-demand. Use it after changing skill docs, tool
    descriptions, CLI help, or other LLM-facing behavior.

.EXAMPLE
    .\scripts\Test-LlmRegressionGate.ps1

.EXAMPLE
    .\scripts\Test-LlmRegressionGate.ps1 -SkipSync
#>

[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$SkipSync,
    [switch]$CliOnly,
    [switch]$McpOnly
)

$ErrorActionPreference = 'Stop'

if ($CliOnly -and $McpOnly) {
    Write-Error "Choose at most one of -CliOnly or -McpOnly."
    exit 1
}

if (-not $env:AZURE_OPENAI_ENDPOINT) {
    Write-Error "AZURE_OPENAI_ENDPOINT is required for the LLM regression gate."
    exit 1
}

$rootDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$llmTestsDir = Join-Path $rootDir "llm-tests"

function Find-BuildArtifact {
    param(
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        $fullPath = Join-Path $rootDir $candidate
        if (Test-Path $fullPath) {
            return (Resolve-Path $fullPath).Path
        }
    }

    return $null
}

if (-not $SkipBuild) {
    Write-Host "Building CLI (Release)..." -ForegroundColor Cyan
    dotnet build (Join-Path $rootDir "src\PptMcp.CLI\PptMcp.CLI.csproj") -c Release
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Write-Host "Building MCP Server (Release)..." -ForegroundColor Cyan
    dotnet build (Join-Path $rootDir "src\PptMcp.McpServer\PptMcp.McpServer.csproj") -c Release
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$cliPath = Find-BuildArtifact @(
    "src\PptMcp.CLI\bin\Release\net9.0-windows\pptcli.exe",
    "src\PptMcp.CLI\bin\Debug\net9.0-windows\pptcli.exe",
    "src\PptMcp.CLI\bin\Release\net10.0-windows\pptcli.exe",
    "src\PptMcp.CLI\bin\Debug\net10.0-windows\pptcli.exe"
)

$mcpServerPath = Find-BuildArtifact @(
    "src\PptMcp.McpServer\bin\Release\net9.0-windows\PptMcp.McpServer.exe",
    "src\PptMcp.McpServer\bin\Debug\net9.0-windows\PptMcp.McpServer.exe",
    "src\PptMcp.McpServer\bin\Release\net10.0-windows\PptMcp.McpServer.exe",
    "src\PptMcp.McpServer\bin\Debug\net10.0-windows\PptMcp.McpServer.exe"
)

if (-not $cliPath) {
    Write-Error "Could not find pptcli.exe. Build src\PptMcp.CLI first."
    exit 1
}

if (-not $mcpServerPath) {
    Write-Error "Could not find PptMcp.McpServer.exe. Build src\PptMcp.McpServer first."
    exit 1
}

$tests = @()

if (-not $McpOnly) {
    $tests += @(
        "cli/test_cli_table.py::test_cli_table_create_query",
        "cli/test_cli_chart.py::test_cli_chart_workflows",
        "cli/test_cli_styling.py::test_cli_styling_header_fill"
    )
}

if (-not $CliOnly) {
    $tests += @(
        "mcp_tests/test_mcp_table.py::test_mcp_table_create_query",
        "mcp_tests/test_mcp_chart.py::test_mcp_chart_workflows",
        "mcp_tests/test_mcp_styling.py::test_mcp_styling_header_fill"
    )
}

$previousCliCommand = [Environment]::GetEnvironmentVariable("PPT_CLI_COMMAND", "Process")
$previousMcpCommand = [Environment]::GetEnvironmentVariable("ppt_mcp_SERVER_COMMAND", "Process")

[Environment]::SetEnvironmentVariable("PPT_CLI_COMMAND", $cliPath, "Process")
[Environment]::SetEnvironmentVariable("ppt_mcp_SERVER_COMMAND", $mcpServerPath, "Process")

Push-Location $llmTestsDir

try {
    if (-not $SkipSync) {
        Write-Host "Syncing llm-tests dependencies..." -ForegroundColor Cyan
        uv sync
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
    }

    Write-Host "Using CLI command: $cliPath" -ForegroundColor Gray
    Write-Host "Using MCP command: $mcpServerPath" -ForegroundColor Gray
    Write-Host "Running canonical LLM regression gate:" -ForegroundColor Cyan
    $tests | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

    uv run pytest -v @tests
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
    [Environment]::SetEnvironmentVariable("PPT_CLI_COMMAND", $previousCliCommand, "Process")
    [Environment]::SetEnvironmentVariable("ppt_mcp_SERVER_COMMAND", $previousMcpCommand, "Process")
}
