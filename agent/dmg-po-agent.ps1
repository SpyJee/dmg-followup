# =============================================================================
# DMG PO PDF Agent
#
# Polls dmg-followup-proxy for PO PDF requests, walks T:\Année En Cours\
# Bon de Commande, uploads to dmgpdfvault, reports back. Outbound HTTPS only.
# Designed for unattended operation on PC11 via Task Scheduler at boot.
#
# Config:
#   $InstallDir\config.json     — non-secret settings (api URL, site, paths)
#   $InstallDir\credentials.dat — DPAPI-encrypted JSON {username, password, site}
#   $InstallDir\logs\           — daily rotating log files
#
# Lifecycle:
#   1. Load config + decrypt credentials
#   2. Login to /auth-login → get bearer token
#   3. Loop:
#       - GET /po-agent-poll?site=DMG
#       - If empty: sleep, continue
#       - Else: walk T:, find PO, request write SAS, upload, complete
#   4. On 401, re-login. On crash, exit (Task Scheduler restarts).
# =============================================================================

param(
    [string]$InstallDir = "$env:ProgramData\DMG\po-agent"
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ----- Config + creds ---------------------------------------------------------

$configPath = Join-Path $InstallDir 'config.json'
$credsPath = Join-Path $InstallDir 'credentials.dat'
$logDir = Join-Path $InstallDir 'logs'

if (-not (Test-Path $configPath)) { throw "config.json not found at $configPath. Run install-agent.ps1 first." }
if (-not (Test-Path $credsPath)) { throw "credentials.dat not found at $credsPath. Run install-agent.ps1 first." }
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$cfg = Get-Content $configPath -Raw | ConvertFrom-Json

# ----- Logging ---------------------------------------------------------------

function Write-AgentLog {
    param([string]$Level, [string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "$ts [$Level] $Message"
    $logFile = Join-Path $logDir ("agent-" + (Get-Date -Format 'yyyy-MM-dd') + ".log")
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch {}
    if ($Level -eq 'ERROR') { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'WARN') { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

# Trim logs older than 14 days at startup.
Get-ChildItem $logDir -Filter "agent-*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-14) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

# ----- DPAPI decrypt creds ---------------------------------------------------

function Read-Credentials {
    param([string]$Path)
    Add-Type -AssemblyName System.Security
    $cipher = [IO.File]::ReadAllBytes($Path)
    $plain = [Security.Cryptography.ProtectedData]::Unprotect(
        $cipher, $null, [Security.Cryptography.DataProtectionScope]::LocalMachine)
    $json = [Text.Encoding]::UTF8.GetString($plain)
    return $json | ConvertFrom-Json
}

# ----- Auth -----------------------------------------------------------------

$script:Token = $null
$script:TokenExpiresAt = [DateTime]::MinValue

function Refresh-Token {
    $creds = Read-Credentials $credsPath
    $body = @{ username = $creds.username; password = $creds.password; site = $creds.site } | ConvertTo-Json -Compress
    Write-AgentLog INFO "Logging in as $($creds.username) at $($cfg.authBase)"
    $r = Invoke-RestMethod -Uri "$($cfg.authBase)/auth-login" -Method Post -ContentType 'application/json' -Body $body -TimeoutSec 30
    if (-not $r.token) { throw "auth-login returned no token" }
    $script:Token = $r.token
    $script:TokenExpiresAt = if ($r.expiresAt) { [DateTime]::Parse($r.expiresAt) } else { (Get-Date).AddHours(8) }
    Write-AgentLog INFO "Token acquired, expires $($script:TokenExpiresAt.ToString('o'))"
}

function Get-AuthHeaders {
    if (-not $script:Token -or (Get-Date) -gt $script:TokenExpiresAt.AddMinutes(-5)) {
        Refresh-Token
    }
    return @{ Authorization = "Bearer $($script:Token)" }
}

# ----- API calls (with single 401 retry) ------------------------------------

function Invoke-Api {
    param(
        [string]$Method = 'GET',
        [string]$Url,
        [object]$Body = $null
    )
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            $headers = Get-AuthHeaders
            $params = @{
                Uri = $Url; Method = $Method; Headers = $headers; TimeoutSec = 60
            }
            if ($Body) {
                $params['ContentType'] = 'application/json'
                $params['Body'] = ($Body | ConvertTo-Json -Compress -Depth 6)
            }
            return Invoke-RestMethod @params
        } catch [System.Net.WebException] {
            $resp = $_.Exception.Response
            if ($resp -and $resp.StatusCode.value__ -eq 401 -and $attempt -eq 1) {
                Write-AgentLog WARN "401 from $Url — re-logging in"
                $script:Token = $null
                continue
            }
            throw
        }
    }
}

# ----- PO file walk (mirrors dmgpo-launcher.ps1 logic) ----------------------

function Find-PoFile {
    param([string]$Po, [string]$PoRoot)

    if (-not (Test-Path $PoRoot)) {
        throw "PO root not reachable: $PoRoot (T: drive mapped?)"
    }

    # Walk every customer folder under the PO root, find any folder/file
    # whose name starts with the PO number. First hit wins.
    $hits = @()
    Get-ChildItem -Path $PoRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $customerFolder = $_
        $matches = Get-ChildItem -Path $customerFolder.FullName -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "$Po*" }
        foreach ($m in $matches) {
            $hits += [pscustomobject]@{
                Customer = $customerFolder.Name
                Item = $m
            }
        }
    }

    if ($hits.Count -eq 0) {
        return $null
    }

    $hit = $hits[0]
    $customer = $hit.Customer
    $item = $hit.Item

    if (-not $item.PSIsContainer) {
        # Direct PDF: e.g., AEROSPHÈRE\503871.pdf
        return [pscustomobject]@{
            Customer = $customer; FilePath = $item.FullName; FileName = $item.Name
        }
    }

    # PO is a folder: pick the best PDF inside.
    $pdfs = Get-ChildItem -Path $item.FullName -Filter "PO $Po*.pdf" -File -ErrorAction SilentlyContinue
    if (-not $pdfs) { $pdfs = Get-ChildItem -Path $item.FullName -Filter "*$Po*.pdf" -File -ErrorAction SilentlyContinue }
    if (-not $pdfs) { $pdfs = Get-ChildItem -Path $item.FullName -Filter "*.pdf" -File -ErrorAction SilentlyContinue }
    if (-not $pdfs) { return $null }

    # Customer path now includes the PO subfolder (mirrors T: layout).
    $rel = "$customer/$($item.Name)"
    return [pscustomobject]@{
        Customer = $rel; FilePath = $pdfs[0].FullName; FileName = $pdfs[0].Name
    }
}

# ----- Blob upload via SAS ---------------------------------------------------

function Upload-ToSas {
    param([string]$SasUrl, [string]$FilePath)
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    # Invoke-WebRequest -UseBasicParsing handles binary bodies cleanly in PS 5.1.
    # Invoke-RestMethod has historically corrupted bytes on PUT.
    $headers = @{
        'x-ms-blob-type' = 'BlockBlob'
        'x-ms-version' = '2020-08-04'
    }
    $resp = Invoke-WebRequest -Method Put -Uri $SasUrl -Headers $headers `
        -Body $bytes -ContentType 'application/pdf' `
        -UseBasicParsing -TimeoutSec 120
    if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
        throw "Upload failed: HTTP $($resp.StatusCode)"
    }
}

# ----- Process one request ---------------------------------------------------

function Process-Request {
    param([object]$Req)
    $id = $Req.id
    $po = $Req.po
    $site = $Req.site
    Write-AgentLog INFO "[req $id] PO=$po site=$site"

    try {
        $found = Find-PoFile -Po $po -PoRoot $cfg.poRoot
        if (-not $found) {
            Write-AgentLog WARN "[req $id] PO file not found on T:"
            Invoke-Api -Method POST -Url "$($cfg.apiBase)/po-agent-complete" -Body @{
                id = $id; site = $site; ok = $false; error = "PO not found on T:"
            } | Out-Null
            return
        }

        $blobPath = "$site/$($found.Customer)/$($found.FileName)"
        Write-AgentLog INFO "[req $id] Found: $($found.FilePath) -> $blobPath"

        # Build the SAS URL piecewise (PS 5.1 won't parse `&` inside an interpolated
        # string used as a parameter value — extract first, then pass).
        $encId = [uri]::EscapeDataString($id)
        $encPath = [uri]::EscapeDataString($blobPath)
        $sasUrl = "$($cfg.apiBase)/po-agent-write-sas?id=$encId" + "&site=$site" + "&blobPath=$encPath"
        $sasResp = Invoke-Api -Method GET -Url $sasUrl
        if (-not $sasResp.writeSas) { throw "no writeSas in response" }

        Upload-ToSas -SasUrl $sasResp.writeSas -FilePath $found.FilePath
        Write-AgentLog INFO "[req $id] Uploaded $((Get-Item $found.FilePath).Length) bytes"

        Invoke-Api -Method POST -Url "$($cfg.apiBase)/po-agent-complete" -Body @{
            id = $id; site = $site; ok = $true; blobPath = $blobPath
        } | Out-Null
        Write-AgentLog INFO "[req $id] DONE"
    } catch {
        $errMsg = $_.Exception.Message
        Write-AgentLog ERROR "[req $id] failed: $errMsg"
        try {
            Invoke-Api -Method POST -Url "$($cfg.apiBase)/po-agent-complete" -Body @{
                id = $id; site = $site; ok = $false; error = $errMsg.Substring(0, [Math]::Min(400, $errMsg.Length))
            } | Out-Null
        } catch { Write-AgentLog ERROR "[req $id] also failed to report failure: $($_.Exception.Message)" }
    }
}

# ----- Main poll loop --------------------------------------------------------

Write-AgentLog INFO "Agent starting. site=$($cfg.site) apiBase=$($cfg.apiBase) poRoot=$($cfg.poRoot) interval=$($cfg.pollIntervalSec)s"

# Heartbeat counter so we don't spam the log every poll.
$idleLogEvery = 300  # log every ~15min when idle
$idleCount = 0

while ($true) {
    try {
        $resp = Invoke-Api -Method GET -Url "$($cfg.apiBase)/po-agent-poll?site=$($cfg.site)"
        if ($resp.empty) {
            $idleCount++
            if ($idleCount -ge $idleLogEvery) {
                Write-AgentLog INFO "Idle heartbeat: no requests in last $($idleCount * $cfg.pollIntervalSec / 60) min"
                $idleCount = 0
            }
        } else {
            $idleCount = 0
            Process-Request -Req $resp
        }
    } catch {
        Write-AgentLog ERROR "Poll loop error: $($_.Exception.Message)"
        # Exponential backoff up to 30s before next poll.
        Start-Sleep -Seconds ([Math]::Min(30, $cfg.pollIntervalSec * 5))
        continue
    }
    Start-Sleep -Seconds $cfg.pollIntervalSec
}
