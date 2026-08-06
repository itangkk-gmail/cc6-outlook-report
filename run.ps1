#Requires -Version 5.1
<#
.SYNOPSIS
  Run CC6 Outlook report pipeline with visible success/failure status.

.EXAMPLE
  .\run.ps1
  .\run.ps1 -NoPause
  .\run.ps1 -MergeOnly -File "sample\No.165Daily Report  2026-6-4.xlsx"
#>
param(
    [switch]$NoPause,
    [switch]$MergeOnly,
    [switch]$Update,
    [string[]]$File = @(),
    [string]$Config = ""
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
New-Item -ItemType Directory -Force -Path "logs" | Out-Null

$start = Get-Date
Write-Host "============================================================"
Write-Host " CC6 Outlook Report"
Write-Host " Start: $($start.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host " Log dir: $(Join-Path $PWD 'logs')"
Write-Host "============================================================"
Write-Host ""

$pythonExe = $null
$python = Get-Command python -ErrorAction SilentlyContinue
if ($python) {
    $pythonExe = "python"
} else {
    $candidate = "$env:USERPROFILE\.workbuddy\binaries\python\versions\3.14.3\python.exe"
    if (Test-Path $candidate) {
        $pythonExe = $candidate
    }
}
if (-not $pythonExe) {
    $msg = "[FAILED] Python not found. Please install Python 3 and add it to PATH."
    Write-Host $msg -ForegroundColor Red
    Set-Content -Path "logs\last_status.txt" -Value @"
status=FAILED
message=Python not found
finished_at=$($start.ToString('yyyy-MM-dd HH:mm:ss'))
"@ -Encoding UTF8
    if (-not $NoPause) { Read-Host "Press Enter to close" | Out-Null }
    exit 1
}
Write-Host "Using Python: $pythonExe"

$argsList = @("main.py")
if ($Config) { $argsList += @("--config", $Config) }
if ($MergeOnly) { $argsList += "--merge-only" }
if ($Update) { $argsList += "--update" }
foreach ($f in $File) { $argsList += @("--file", $f) }

try {
    & $pythonExe @argsList
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host "[FAILED] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
}

Write-Host ""
Write-Host "============================================================"
if ($exitCode -ne 0) {
    Write-Host " RESULT: FAILED  exit_code=$exitCode" -ForegroundColor Red
} else {
    Write-Host " RESULT: SUCCESS" -ForegroundColor Green
}
Write-Host " See: logs\last_status.txt"
Write-Host " See: logs\run.log and latest logs\run_*.log"
Write-Host "============================================================"

if (-not $NoPause) {
    Read-Host "Press Enter to close" | Out-Null
}
exit $exitCode
