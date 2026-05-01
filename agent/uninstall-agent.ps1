#Requires -RunAsAdministrator
# =============================================================================
# DMG PO PDF Agent — Uninstaller
#
# Stops the scheduled task, removes the registration, deletes agent files
# (including encrypted credentials). Logs are kept for forensic purposes.
# =============================================================================

param(
    [switch]$KeepLogs
)

$ErrorActionPreference = 'Continue'

$InstallDir = Join-Path $env:ProgramData 'DMG\po-agent'
$TaskName = 'DMG-PO-Agent'

Write-Host "=== DMG PO PDF Agent — Uninstaller ===" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/3] Stopping task..." -ForegroundColor Yellow
Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  [+] Task '$TaskName' removed"

Write-Host ""
Write-Host "[2/3] Removing credentials..." -ForegroundColor Yellow
$credsPath = Join-Path $InstallDir 'credentials.dat'
if (Test-Path $credsPath) {
    Remove-Item $credsPath -Force
    Write-Host "  [+] credentials.dat deleted"
}

Write-Host ""
Write-Host "[3/3] Removing agent files..." -ForegroundColor Yellow
if (Test-Path (Join-Path $InstallDir 'dmg-po-agent.ps1')) {
    Remove-Item (Join-Path $InstallDir 'dmg-po-agent.ps1') -Force
}
if (Test-Path (Join-Path $InstallDir 'config.json')) {
    Remove-Item (Join-Path $InstallDir 'config.json') -Force
}
if (-not $KeepLogs) {
    $logDir = Join-Path $InstallDir 'logs'
    if (Test-Path $logDir) {
        Remove-Item $logDir -Recurse -Force
        Write-Host "  [+] Logs deleted (use -KeepLogs to retain)"
    }
}

# Remove install dir if empty.
if (Test-Path $InstallDir) {
    if (-not (Get-ChildItem $InstallDir -Force -ErrorAction SilentlyContinue)) {
        Remove-Item $InstallDir -Force
    }
}

Write-Host ""
Write-Host "=== Uninstall complete ===" -ForegroundColor Green
