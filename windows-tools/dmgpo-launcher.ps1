# DMG PO URL handler — opens the PO PDF or Explorer for a given PO number.
# Invoked by Windows when a user clicks dmgpo://<PO> in a browser.
#
# Argument: full URL like "dmgpo://5000024373" or "dmgpo://5000024373/"
#
# Search rules (case-insensitive):
#   1. Walk every customer folder under the PO root.
#   2. Find any folder OR file whose name starts with the PO number, looking
#      one level deep (covers AEROSPHÈRE/<po>.pdf style) AND two levels deep
#      (covers AIRBUS/<5-digit-prefix>/<full-po>/ style).
#   3. If the match is a folder, look inside for "PO <number>*.pdf"
#      (or any *.pdf containing the PO number) and open the first match.
#   4. If the match is a file, open it directly.
#   5. If nothing matches, open the PO root so the user can browse.

param([string]$Url)

$ErrorActionPreference = 'Stop'
$PoRoot = 'T:\Année En Cours\Bon de Commande'

# Load assemblies up front. System.Windows.Forms isn't auto-loaded in PS 5.1
# and we need it for the "T: not mapped" message box.
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

# Parse the PO out of the URL: strip scheme, trailing slash, URL-decode.
$po = $Url -replace '^[a-z]+://', '' -replace '/+$', ''
$po = [System.Web.HttpUtility]::UrlDecode($po)
$po = $po.Trim()

if ([string]::IsNullOrWhiteSpace($po)) {
    Start-Process explorer.exe -ArgumentList "`"$PoRoot`""
    exit 0
}

if (-not (Test-Path $PoRoot)) {
    [System.Windows.Forms.MessageBox]::Show(
        "PO folder not reachable:`n$PoRoot`n`nIs the T: drive mapped?",
        "DMG PO Opener", 0, 16) | Out-Null
    exit 1
}

# Find any folder/file across all customer subdirectories whose name starts
# with the PO number. Most POs are unique across customers, so the first hit
# wins. If no match, fall back to the PO root.
#
# Why two passes:
#   AEROSPHÈRE puts the PDF directly in the customer folder (depth 0).
#   AIRBUS / Bombardier US group POs by 3-5 digit prefix subfolders, so the
#   PO file/folder lives one level deeper (depth 1). Cheap pass at depth 0
#   first, fall back to depth 1 only if nothing found — avoids scanning
#   thousands of items in customer folders that already had a hit.
$hits = @()
Get-ChildItem -Path $PoRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $hits += Get-ChildItem -Path $_.FullName -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$po*" }
}
if ($hits.Count -eq 0) {
    Get-ChildItem -Path $PoRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $hits += Get-ChildItem -Path $_.FullName -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "$po*" }
        }
    }
}

if ($hits.Count -eq 0) {
    Start-Process explorer.exe -ArgumentList "`"$PoRoot`""
    exit 0
}

$first = $hits[0]
if ($first.PSIsContainer) {
    # PO is a folder — look for a PDF inside, prefer "PO <num>*.pdf"
    $pdfs = Get-ChildItem -Path $first.FullName -Filter "PO $po*.pdf" -ErrorAction SilentlyContinue
    if (-not $pdfs) {
        $pdfs = Get-ChildItem -Path $first.FullName -Filter "*$po*.pdf" -ErrorAction SilentlyContinue
    }
    if (-not $pdfs) {
        $pdfs = Get-ChildItem -Path $first.FullName -Filter "*.pdf" -ErrorAction SilentlyContinue
    }
    if ($pdfs) {
        Start-Process $pdfs[0].FullName
    } else {
        Start-Process explorer.exe -ArgumentList "`"$($first.FullName)`""
    }
} else {
    Start-Process $first.FullName
}
