<#
.SYNOPSIS
    Stops the PptMcp Service gracefully and kills PowerPoint processes before build.
.DESCRIPTION
    Pre-build cleanup script that:
    1. Gracefully stops the PptMcp Service via named pipe (service.shutdown)
    2. Kills any remaining PowerPoint (POWERPNT.EXE) processes

    This prevents file locking issues during build when the service or PowerPoint
    holds handles to assemblies or presentations.
.NOTES
    Called from Directory.Build.props as a BeforeBuild target.
    Safe to run when no processes are running (silently succeeds).
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = 'SilentlyContinue'

function Write-Status($message) {
    if ($Verbose) {
        Write-Host "  [pre-build] $message" -ForegroundColor DarkGray
    }
}

# ----------------------------------------------
# 1. Gracefully stop PptMcp Service via CLI
# ----------------------------------------------
function Stop-PptMcpService {
    # Look for pptcli in build output directories (Debug/Release)
    $scriptDir = Split-Path -Parent $PSScriptRoot  # repo root
    $cliPaths = @(
        "$scriptDir\src\PptMcp.CLI\bin\Debug\net10.0-windows\pptcli.exe",
        "$scriptDir\src\PptMcp.CLI\bin\Release\net10.0-windows\pptcli.exe"
    )
    $pptcli = $cliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($pptcli) {
        Write-Status "Using CLI: $pptcli"
        $output = & $pptcli service stop --quiet 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            # Parse JSON to check if service was running
            try {
                $result = $output | ConvertFrom-Json
                if ($result.message -eq 'Service is not running.') {
                    Write-Status "PptMcp Service was not running"
                } else {
                    Write-Host "  PptMcp Service stopped gracefully" -ForegroundColor Green
                }
            } catch {
                Write-Status "Service stop completed (exit code 0)"
            }
        } else {
            Write-Status "CLI service stop returned exit code $exitCode, falling back to process kill"
            Stop-PptMcpServiceFallback
        }
    } else {
        Write-Status "pptcli not found (first build?), using fallback"
        Stop-PptMcpServiceFallback
    }
}

function Stop-PptMcpServiceFallback {
    # Fallback: direct named pipe shutdown (works without CLI binary)
    $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    $pipeName = "PptMcp-$sid"

    $pipeExists = Test-Path "\\.\pipe\$pipeName"
    if (-not $pipeExists) {
        Write-Status "PptMcp Service not running (no pipe found)"
        return
    }

    Write-Status "PptMcp Service detected, sending shutdown via pipe..."
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect(3000)

        $writer = New-Object System.IO.StreamWriter($pipe, [System.Text.Encoding]::UTF8, 4096)
        $writer.AutoFlush = $true
        $reader = New-Object System.IO.StreamReader($pipe, [System.Text.Encoding]::UTF8)

        $writer.WriteLine('{"Command":"service.shutdown"}')
        $response = $reader.ReadLine()
        Write-Status "Service response: $response"

        $reader.Dispose()
        $writer.Dispose()
        $pipe.Dispose()

        Start-Sleep -Milliseconds 500
        Write-Host "  PptMcp Service stopped gracefully" -ForegroundColor Green
    }
    catch {
        Write-Status "Could not connect to pipe: $($_.Exception.Message)"
        $serviceProcs = Get-Process -Name 'PptMcp.McpServer', 'PptMcp.Service' -ErrorAction SilentlyContinue
        if ($serviceProcs) {
            $serviceProcs | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host "  PptMcp Service processes killed (pipe unavailable)" -ForegroundColor Yellow
        }
    }
}

# ----------------------------------------------
# 2. Kill PowerPoint processes
# ----------------------------------------------
function Stop-PowerPointProcesses {
    $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
    if ($excelProcs) {
        $count = $excelProcs.Count
        $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        Write-Host "  Killed $count PowerPoint process(es)" -ForegroundColor Yellow
    }
    else {
        Write-Status "No PowerPoint processes running"
    }
}

# ----------------------------------------------
# Run cleanup
# ----------------------------------------------
Stop-PptMcpService
Stop-PowerPointProcesses
