#Requires -RunAsAdministrator
# =============================================================================
# DMG PO PDF Agent — Installer
#
# Run as Administrator on the on-prem machine that should host the agent
# (PC11 for DMG today; a DICI box later).
#
# RUNTIME MODEL: the agent runs as the installing user (so it inherits the
# user's Kerberos rights to the file share + DPAPI key) AT SYSTEM STARTUP,
# even when nobody is logged in. To make this work, Task Scheduler stores
# the user's Windows password in the LSA secrets store. If the Windows
# password ever changes, re-run this installer.
#
# WHY UNC INSTEAD OF T:\\: drive-letter mappings (T:) only exist inside an
# interactive logon session, so a task running pre-login can't see T:\.
# The agent uses the UNC path \\\\WIN2K19\\BaseDMG\\... directly.
#
# Steps performed:
#   1. Verify file share is reachable
#   2. Prompt for DMG portal credentials, test login
#   3. Prompt for the Windows password of the run-as account
#   4. DPAPI-encrypt portal credentials (CurrentUser scope)
#   5. Copy agent files to %ProgramData%\\DMG\\po-agent
#   6. Write config.json
#   7. Register a Scheduled Task: AtStartup, run as <user>, password saved
#
# Re-running is safe (idempotent): config + task are overwritten.
# =============================================================================

param(
    [string]$Site = 'DMG',
    [string]$PoRoot = '\\WIN2K19\BaseDMG\Année En Cours\Bon de Commande',
    [string]$ApiBase = 'https://dmg-followup-proxy.azurewebsites.net/api/followup',
    [string]$AuthBase = 'https://dmg-proxy-b4fpgjh8begye3fb.canadacentral-01.azurewebsites.net/api/dmg',
    [int]$PollIntervalSec = 3,
    [string]$RunAsUser = "$env:USERDOMAIN\$env:USERNAME"
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$InstallDir = Join-Path $env:ProgramData 'DMG\po-agent'
$TaskName = 'DMG-PO-Agent'

Write-Host "=== DMG PO PDF Agent — Installer ===" -ForegroundColor Cyan
Write-Host "Install dir: $InstallDir"
Write-Host "Site:        $Site"
Write-Host "PO root:     $PoRoot"
Write-Host "API base:    $ApiBase"
Write-Host "Run-as user: $RunAsUser"
Write-Host ""

# ----- 1. Verify T: drive ----------------------------------------------------

Write-Host "[1/7] Checking PO root..." -ForegroundColor Yellow
if (-not (Test-Path $PoRoot)) {
    Write-Host "  [X] PO root not reachable: $PoRoot" -ForegroundColor Red
    Write-Host "      Make sure the file share is reachable from this account."
    exit 1
}
$customers = (Get-ChildItem -Path $PoRoot -Directory -ErrorAction SilentlyContinue).Count
Write-Host "  [+] PO root OK ($customers customer folders)"

# ----- 2. Prompt for credentials + test login --------------------------------

Write-Host ""
Write-Host "[2/7] DMG portal credentials..." -ForegroundColor Yellow
Write-Host "      The agent will use these to authenticate to the proxy."
Write-Host "      Recommendation: create a dedicated 'po-agent' user account."
Write-Host ""

$user = Read-Host "DMG username"
$secPass = Read-Host "DMG password" -AsSecureString
$plainPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPass))

Write-Host ""
Write-Host "  Testing login..." -NoNewline
$loginBody = @{ username = $user; password = $plainPass; site = $Site } | ConvertTo-Json -Compress
try {
    $loginRes = Invoke-RestMethod -Uri "$AuthBase/auth-login" -Method Post -ContentType 'application/json' -Body $loginBody -TimeoutSec 30
    if (-not $loginRes.token) { throw "auth-login returned no token" }
    Write-Host " OK" -ForegroundColor Green
    Write-Host "  Token expires: $($loginRes.expiresAt)"
} catch {
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)"
    exit 1
}

# ----- 3. DPAPI-encrypt credentials ------------------------------------------

Write-Host ""
Write-Host "[3/7] Windows password for $RunAsUser..." -ForegroundColor Yellow
Write-Host "      Task Scheduler will store this so the agent can run pre-login (after a reboot)."
Write-Host "      Stored in the Windows LSA secrets store, not in any file. Re-run the installer if"
Write-Host "      this account's Windows password ever changes."
Write-Host ""
$secWinPass = Read-Host "Windows password for $RunAsUser" -AsSecureString
$plainWinPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secWinPass))

Write-Host ""
Write-Host "[4/7] Encrypting portal credentials..." -ForegroundColor Yellow

if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }

Add-Type -AssemblyName System.Security
$credObj = @{ username = $user; password = $plainPass; site = $Site }
$plainBytes = [Text.Encoding]::UTF8.GetBytes(($credObj | ConvertTo-Json -Compress))
# CurrentUser scope: only the installing user can decrypt. The agent runs
# as the same user, so it can decrypt; SYSTEM, other users, or anyone with
# admin-but-different-account cannot. Stronger than LocalMachine scope.
$cipher = [Security.Cryptography.ProtectedData]::Protect(
    $plainBytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
$credsPath = Join-Path $InstallDir 'credentials.dat'
[IO.File]::WriteAllBytes($credsPath, $cipher)
# Tighten ACL: only SYSTEM, Administrators, and the installing user.
$me = "$env:USERDOMAIN\$env:USERNAME"
$acl = Get-Acl $credsPath
$acl.SetAccessRuleProtection($true, $false)  # disable inheritance, drop existing rules
$rules = @(
    (New-Object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM",   "FullControl", "Allow")),
    (New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")),
    (New-Object System.Security.AccessControl.FileSystemAccessRule($me,                      "FullControl", "Allow"))
)
foreach ($r in $rules) { $acl.AddAccessRule($r) }
Set-Acl -Path $credsPath -AclObject $acl
Write-Host "  [+] credentials.dat written + ACLed (SYSTEM + Admins + $me)"

# Zero out the plaintext password from memory (best-effort).
$plainPass = $null
[GC]::Collect()

# ----- 4. Copy agent files ---------------------------------------------------

Write-Host ""
Write-Host "[5/7] Copying agent files..." -ForegroundColor Yellow
$here = Split-Path -Parent $PSCommandPath
$src = Join-Path $here 'dmg-po-agent.ps1'
if (-not (Test-Path $src)) {
    Write-Host "  [X] Cannot find dmg-po-agent.ps1 next to this installer." -ForegroundColor Red
    exit 1
}
Copy-Item $src -Destination (Join-Path $InstallDir 'dmg-po-agent.ps1') -Force
Write-Host "  [+] dmg-po-agent.ps1 -> $InstallDir"

# ----- 5. Write config.json --------------------------------------------------

Write-Host ""
Write-Host "[6/7] Writing config..." -ForegroundColor Yellow
$config = @{
    apiBase = $ApiBase
    authBase = $AuthBase
    site = $Site
    poRoot = $PoRoot
    pollIntervalSec = $PollIntervalSec
}
$configJson = $config | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText((Join-Path $InstallDir 'config.json'), $configJson, [Text.UTF8Encoding]::new($true))
Write-Host "  [+] config.json written"

# ----- 6. Register Scheduled Task -------------------------------------------

Write-Host ""
Write-Host "[7/7] Registering Scheduled Task..." -ForegroundColor Yellow

# Remove existing task if present.
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\dmg-po-agent.ps1`""

# AtStartup so the agent runs even when nobody is logged in (per JP: PC11 stays on
# 24/7 but you have to enter your password to start a Windows session). Also a
# one-shot now-trigger so the agent starts without waiting for the next reboot.
$triggerStartup = New-ScheduledTaskTrigger -AtStartup
$triggerNow = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(30)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 9999 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Days 0) `
    -MultipleInstances IgnoreNew

# Pass -User and -Password directly to Register-ScheduledTask. This sets logonType
# to Password (saves the password in the LSA secret store, encrypted, accessible
# only to the SYSTEM/Task Scheduler service). With this, the task can launch as
# $RunAsUser even pre-login.
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($triggerStartup, $triggerNow) `
    -User $RunAsUser `
    -Password $plainWinPass `
    -RunLevel Highest `
    -Settings $settings `
    -Description "DMG PO PDF agent — polls dmg-followup-proxy and uploads PO PDFs on demand." | Out-Null

# Zero out the plaintext password from memory (best-effort).
$plainWinPass = $null
[GC]::Collect()

Write-Host "  [+] Task '$TaskName' registered"
Write-Host "      Runs as $RunAsUser, starts at boot (pre-login), restarts every 1 min on failure"

Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

$task = Get-ScheduledTask -TaskName $TaskName
$taskInfo = $task | Get-ScheduledTaskInfo
Write-Host ""
Write-Host "Task state: $($task.State)"
Write-Host "Last run:   $($taskInfo.LastRunTime)"
Write-Host "Last result: $($taskInfo.LastTaskResult)  (267009 = currently running)"

Write-Host ""
Write-Host "=== Install complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Logs:  $InstallDir\logs\agent-YYYY-MM-DD.log"
Write-Host "Stop:  Stop-ScheduledTask -TaskName '$TaskName'"
Write-Host "Start: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Tail:  Get-Content `"$InstallDir\logs\agent-$(Get-Date -Format 'yyyy-MM-dd').log`" -Wait -Tail 20"
