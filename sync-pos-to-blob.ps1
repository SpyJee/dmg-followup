# =============================================================================
# DMG PO PDF sync — T:\Bon de Commande -> dmgpdfvault/po-pdfs
#
# Walks every customer folder under T:\Année En Cours\Bon de Commande, finds
# every PDF (flat layouts like AEROSPHÈRE\<po>.pdf and prefix-nested layouts
# like AIRBUS\<group>\<po>\*.pdf), and uploads new/changed files to the
# dmgpdfvault container. Idempotent — only uploads files whose size or
# LastWriteTime differs from the existing blob. Safe to run hourly via
# Task Scheduler.
#
# Auth: pulls the storage account key from `az storage account keys list`
# (requires `az login` to have run interactively at least once on this box).
# Set the env var DMG_BLOB_KEY to skip the az lookup if running unattended.
#
# Usage:
#   .\sync-pos-to-blob.ps1                     # default: full T: walk
#   .\sync-pos-to-blob.ps1 -CustomerFilter AEROSPHÈRE  # one customer only
#   .\sync-pos-to-blob.ps1 -DryRun             # don't upload, just report
# =============================================================================

param(
    [string]$PoRoot = 'T:\Année En Cours\Bon de Commande',
    [string]$Account = 'dmgpdfvault',
    [string]$Container = 'po-pdfs',
    [string]$Site = 'DMG',
    [string]$CustomerFilter = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ----- Logging --------------------------------------------------------------

$logDir = "$env:LOCALAPPDATA\DMG\sync-logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logFile = Join-Path $logDir ("sync-pos-" + (Get-Date -Format 'yyyy-MM-dd') + ".log")
function Log {
    param([string]$Level, [string]$Msg)
    $line = "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg"
    try { Add-Content $logFile $line -Encoding UTF8 } catch {}
    if ($Level -eq 'ERR')   { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'UP'){ Write-Host $line -ForegroundColor Green }
    elseif ($Level -eq 'WARN'){ Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

# Trim logs older than 30 days.
Get-ChildItem $logDir -Filter "sync-pos-*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

Log INFO "=== Sync start: T: -> $Account/$Container (site=$Site) ==="
if ($DryRun) { Log WARN "Dry run: no uploads will happen" }

# ----- Get storage account key ----------------------------------------------

$Key = $env:DMG_BLOB_KEY
if (-not $Key) {
    Log INFO "Getting account key from az CLI..."
    $Key = az storage account keys list -g dmg-quoting-rg -n $Account --query "[0].value" -o tsv 2>$null
    if (-not $Key) {
        Log ERR "Cannot get storage key. Run 'az login' or set DMG_BLOB_KEY env var."
        exit 1
    }
}
$KeyBytes = [Convert]::FromBase64String($Key)

# ----- Blob REST API helpers ------------------------------------------------

# Build a SharedKey-signed request and return the response body (string).
function Invoke-BlobReq {
    param(
        [string]$Method,
        [string]$PathWithQuery,
        [byte[]]$Body = $null,
        [string]$ContentType = '',
        [hashtable]$ExtraMsHeaders = @{}
    )
    $date = (Get-Date).ToUniversalTime().ToString('R')
    $version = '2020-08-04'
    $contentLength = if ($Body) { $Body.Length.ToString() } else { '' }

    # Canonicalized resource: /<account>/<path>\n<sorted query params>
    $parts = $PathWithQuery -split '\?', 2
    $pathOnly = $parts[0]
    $qs = if ($parts.Length -gt 1) { $parts[1] } else { '' }
    $cr = "/$Account$pathOnly"
    if ($qs) {
        $params = @{}
        foreach ($kv in $qs -split '&') {
            $kvp = $kv -split '=', 2
            $k = [System.Web.HttpUtility]::UrlDecode($kvp[0]).ToLower()
            $v = if ($kvp.Length -gt 1) { [System.Web.HttpUtility]::UrlDecode($kvp[1]) } else { '' }
            if ($params.ContainsKey($k)) { $params[$k] = "$($params[$k]),$v" } else { $params[$k] = $v }
        }
        foreach ($k in ($params.Keys | Sort-Object)) {
            $cr += "`n$k`:$($params[$k])"
        }
    }

    # Canonicalized x-ms-* headers (lowercase, sorted, header:value\n)
    $msHeaders = @{ 'x-ms-date' = $date; 'x-ms-version' = $version }
    if ($Body) { $msHeaders['x-ms-blob-type'] = 'BlockBlob' }
    foreach ($k in $ExtraMsHeaders.Keys) { $msHeaders[$k] = $ExtraMsHeaders[$k] }
    $ch = ''
    foreach ($k in ($msHeaders.Keys | Sort-Object)) { $ch += "$($k.ToLower())`:$($msHeaders[$k])`n" }
    $ch = $ch.TrimEnd("`n")

    $stringToSign = ($Method, '', '', $contentLength, '', $ContentType, '', '', '', '', '', '') -join "`n"
    $stringToSign += "`n$ch`n$cr"

    $hmac = New-Object System.Security.Cryptography.HMACSHA256 (,$KeyBytes)
    $sig = [Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($stringToSign)))
    $hmac.Dispose()

    $headers = @{ 'Authorization' = "SharedKey $($Account):$sig" }
    foreach ($k in $msHeaders.Keys) { $headers[$k] = $msHeaders[$k] }
    if ($ContentType) { $headers['Content-Type'] = $ContentType }

    $url = "https://$Account.blob.core.windows.net$PathWithQuery"
    if ($Body) {
        return Invoke-WebRequest -Uri $url -Method $Method -Headers $headers `
            -Body $Body -ContentType $ContentType -UseBasicParsing -TimeoutSec 120
    } else {
        return Invoke-WebRequest -Uri $url -Method $Method -Headers $headers `
            -UseBasicParsing -TimeoutSec 60
    }
}

Add-Type -AssemblyName System.Web

# List all existing blobs (paginated). Returns @{ "<blobPath>" = @{ size; lastModified } }.
function Get-AllBlobs {
    $all = @{}
    $marker = $null
    while ($true) {
        $qs = "restype=container&comp=list&maxresults=5000"
        if ($marker) { $qs += "&marker=$([uri]::EscapeDataString($marker))" }
        $resp = Invoke-BlobReq -Method GET -PathWithQuery "/$Container`?$qs"
        # Azure returns the body with a UTF-8 BOM; PS [xml] chokes regardless of
        # whether the BOM is decoded as FEFF or as 3 Latin-1 chars (ï»¿).
        # Strip everything up to the first '<' to be safe.
        $content = $resp.Content
        $i = $content.IndexOf('<')
        if ($i -gt 0) { $content = $content.Substring($i) }
        $xml = [xml]$content
        foreach ($b in $xml.EnumerationResults.Blobs.Blob) {
            $all[$b.Name] = @{
                size = [long]$b.Properties.'Content-Length'
                lastModified = [DateTime]::Parse($b.Properties.'Last-Modified')
            }
        }
        $next = $xml.EnumerationResults.NextMarker
        if (-not $next) { break }
        $marker = $next
    }
    return $all
}

# Upload a local file as a block blob at $blobPath. URL-encodes path segments.
function Upload-Blob {
    param([string]$LocalPath, [string]$BlobPath)
    $bytes = [IO.File]::ReadAllBytes($LocalPath)
    $encoded = ($BlobPath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    Invoke-BlobReq -Method PUT -PathWithQuery "/$Container/$encoded" `
        -Body $bytes -ContentType 'application/pdf' | Out-Null
}

# ----- Main -----------------------------------------------------------------

if (-not (Test-Path $PoRoot)) {
    Log ERR "PO root not reachable: $PoRoot"
    exit 1
}

Log INFO "Listing existing blobs in $Account/$Container ..."
$existing = Get-AllBlobs
Log INFO "  Found $($existing.Count) existing blobs"

$stats = @{ scanned = 0; uploaded = 0; skipped = 0; failed = 0 }

# Process one PDF: decide upload or skip, then act.
function Sync-OnePdf {
    param([string]$LocalPath, [string]$BlobPath)
    $script:stats.scanned++
    $local = Get-Item $LocalPath
    $cached = $existing[$BlobPath]
    # Skip if same size AND blob's LastModified is newer than local LastWriteTime
    # (blob can't be older than the source it was uploaded from).
    if ($cached -and $cached.size -eq $local.Length -and $cached.lastModified -gt $local.LastWriteTimeUtc) {
        $script:stats.skipped++
        return
    }
    Log UP "UPLOAD $BlobPath ($([math]::Round($local.Length/1KB)) KB)"
    if ($DryRun) { return }
    try {
        Upload-Blob -LocalPath $LocalPath -BlobPath $BlobPath
        $script:stats.uploaded++
    } catch {
        Log ERR "  failed: $($_.Exception.Message)"
        $script:stats.failed++
    }
}

# Walk customer folders (depth 0) — direct PDFs and one-level-down PDFs.
$customers = Get-ChildItem -Path $PoRoot -Directory -ErrorAction SilentlyContinue
if ($CustomerFilter) {
    $customers = $customers | Where-Object { $_.Name -like "*$CustomerFilter*" }
    Log INFO "Filtered to $($customers.Count) customer folder(s) matching '$CustomerFilter'"
}

foreach ($cust in $customers) {
    $custStart = Get-Date
    $custCount = $stats.uploaded + $stats.skipped + $stats.failed

    # Depth 0: PDFs directly under customer folder
    Get-ChildItem $cust.FullName -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
        $blobPath = "$Site/$($cust.Name)/$($_.Name)"
        Sync-OnePdf -LocalPath $_.FullName -BlobPath $blobPath
    }

    # Depth 1: PDFs in subfolders (covers AIRBUS\50023\<po>\*.pdf etc.)
    Get-ChildItem $cust.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $sub = $_
        Get-ChildItem $sub.FullName -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
            $blobPath = "$Site/$($cust.Name)/$($sub.Name)/$($_.Name)"
            Sync-OnePdf -LocalPath $_.FullName -BlobPath $blobPath
        }
        # Depth 2: PDFs nested two levels deep (some customers have group/po/files.pdf)
        Get-ChildItem $sub.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $sub2 = $_
            Get-ChildItem $sub2.FullName -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
                $blobPath = "$Site/$($cust.Name)/$($sub.Name)/$($sub2.Name)/$($_.Name)"
                Sync-OnePdf -LocalPath $_.FullName -BlobPath $blobPath
            }
        }
    }

    $custDur = (Get-Date) - $custStart
    $delta = ($stats.uploaded + $stats.skipped + $stats.failed) - $custCount
    if ($delta -gt 0) {
        Log INFO "  $($cust.Name): $delta files in $([math]::Round($custDur.TotalSeconds, 1))s"
    }
}

Log INFO "=== Sync done: scanned=$($stats.scanned) uploaded=$($stats.uploaded) skipped=$($stats.skipped) failed=$($stats.failed) ==="
