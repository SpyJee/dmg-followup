# ══════════════════════════════════════════════════════════════
# DMG Follow-Up — SQL Server Sync Script
# Reads active orders from WIN2K19\SQLEXPRESS and pushes to Azure API
# Usage: powershell -ExecutionPolicy Bypass -File sync-from-sql.ps1
# ══════════════════════════════════════════════════════════════

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$API_BASE = 'https://dmg-followup-proxy.azurewebsites.net/api/followup'
$SQL_CONN = 'Server=WIN2K19\SQLEXPRESS;Database=DMG_BASE_TAKe2;Trusted_Connection=Yes;'

# Only sync orders with dateRequise >= this date (skip old unarchived records)
$DATE_CUTOFF = '2025-01-01'

# ── Step 1: Connect to SQL Server ──
Write-Host "`n=== DMG Follow-Up SQL Sync ===" -ForegroundColor Cyan
Write-Host "Connecting to SQL Server..." -ForegroundColor Yellow

$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = $SQL_CONN
$conn.Open()
Write-Host "Connected to WIN2K19\SQLEXPRESS" -ForegroundColor Green

# ── Step 2: Get column names (handles encoding) ──
$colCmd = $conn.CreateCommand()
$colCmd.CommandText = 'SELECT TOP 1 * FROM [*REG_ITEM]'
$colAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($colCmd)
$colDt = New-Object System.Data.DataTable
$colAdapter.Fill($colDt) | Out-Null
$numCol = $colDt.Columns[1].ColumnName  # "Numéro" with proper encoding

# Get CLIENTS_FOURNISSEURS société column
$cfCmd = $conn.CreateCommand()
$cfCmd.CommandText = 'SELECT TOP 1 * FROM CLIENTS_FOURNISSEURS'
$cfAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($cfCmd)
$cfDt = New-Object System.Data.DataTable
$cfAdapter.Fill($cfDt) | Out-Null
$socCol = $cfDt.Columns[1].ColumnName  # "SOCIÉTÉ" with proper encoding

# Get LIVRAISON column names
$livCmd = $conn.CreateCommand()
$livCmd.CommandText = 'SELECT TOP 1 * FROM [*REG_LIVRAISON]'
$livAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($livCmd)
$livDt = New-Object System.Data.DataTable
$livAdapter.Fill($livDt) | Out-Null
# Key columns by ordinal:
# [0] POITEMID, [1] Numéro, [3] QTY COMMANDÉE, [4] COMMENTS,
# [5] DATE REQUISE, [6] MG NUMÉRO, [8] QTY LIVRÉ, [9] DATE LIVRÉ,
# [14] cancellé, [15] ARCHIVE, [24] SUIVI
$qtyCol = $livDt.Columns[3].ColumnName      # QTY COMMANDÉE
$dateReqCol = $livDt.Columns[5].ColumnName   # DATE REQUISE
$woCol = $livDt.Columns[6].ColumnName        # MG NUMÉRO
$qtyLivCol = $livDt.Columns[8].ColumnName    # QTY LIVRÉ
$dateLivCol = $livDt.Columns[9].ColumnName   # DATE LIVRÉ
$cancelCol = $livDt.Columns[14].ColumnName   # cancellé
$suiviCol = $livDt.Columns[24].ColumnName    # SUIVI

# ── Step 3: Query active orders ──
Write-Host "Querying active orders..." -ForegroundColor Yellow

$query = @"
SELECT
    p.[BON DE COMMANDE] AS col_po,
    i.item AS col_item,
    i.[PART NUMBER] AS col_partNumber,
    i.NOM AS col_partName,
    cf.[$socCol] AS col_client,
    l.[$qtyCol] AS col_qty,
    l.[$dateReqCol] AS col_dateRequise,
    l.[$qtyLivCol] AS col_qtyLivre,
    l.[$dateLivCol] AS col_dateLivre,
    l.COMMENTS AS col_comment,
    l.[$suiviCol] AS col_suivi,
    l.[$cancelCol] AS col_cancelled,
    l.ARCHIVE AS col_archived,
    l.[$woCol] AS col_workOrder,
    a.NOM_ACHETEUR AS col_buyer
FROM [*REG_PO] p
JOIN [*REG_ITEM] i ON i.po_ident = p.po_ident
JOIN [*REG_LIVRAISON] l ON l.[$numCol] = i.[$numCol]
LEFT JOIN CLIENTS_FOURNISSEURS cf ON cf.FOURNISSEUR_ID = p.FOURNISSEUR_ID
LEFT JOIN ACHETEURS a ON a.ACHETEUR_ID = p.ACHETEUR_ID
WHERE l.[$dateReqCol] IS NOT NULL
  AND ISNULL(l.[$cancelCol], 0) = 0
  AND ISNULL(l.ARCHIVE, 0) = 0
  AND l.[$dateReqCol] >= '$DATE_CUTOFF'
ORDER BY l.[$dateReqCol] ASC
"@

$cmd = $conn.CreateCommand()
$cmd.CommandText = $query
$adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
$count = $adapter.Fill($dt)
$conn.Close()

Write-Host "Found $count active orders in SQL Server" -ForegroundColor Green

# ── Step 4: Map to Follow-Up order format ──
Write-Host "Mapping orders..." -ForegroundColor Yellow

$orders = @()
$seen = @{}

foreach ($row in $dt.Rows) {
    $po = [string]$row['col_po']
    $item = [string]$row['col_item']
    if ([string]::IsNullOrWhiteSpace($po)) { continue }

    $partNumber = if ($row['col_partNumber'] -ne [DBNull]::Value) { [string]$row['col_partNumber'] } else { '' }
    $pnUpper = $partNumber.Trim().ToUpper()
    if ($pnUpper -eq 'FAI' -or $pnUpper -eq 'NRC') { continue }

    $commandeItem = "$po-$item"

    # Skip duplicates (same PO-item, keep first = earliest date)
    if ($seen.ContainsKey($commandeItem)) { continue }
    $seen[$commandeItem] = $true

    # Format date
    $dateRequise = ''
    if ($row['col_dateRequise'] -ne [DBNull]::Value) {
        $dateRequise = ([DateTime]$row['col_dateRequise']).ToString('yyyy-MM-dd')
    }

    # Map client name to match existing client list
    $rawClient = if ($row['col_client'] -ne [DBNull]::Value) { [string]$row['col_client'] } else { '' }
    $client = switch -Wildcard ($rawClient.ToUpper()) {
        '*BOMBARDIER AERO*'   { 'BOMBARDIER INC.' }
        '*BOMBARDIER INC*'    { 'BOMBARDIER INC.' }
        '*BOMBARDIER US*'     { 'BOMBARDIER US AEROSTRUCTURES' }
        '*AIRBUS CANADA*'     { 'AIRBUS CANADA' }  # catches "AIRBUS CANADA LTD PARTNERSHIP" too
        '*AIRBUS ATLANTIC*'   { 'AIRBUS ATLANTIQUE CANADA INC.' }
        '*AIRBUS U*'          { 'AIRBUS U.S A220 INC.' }
        '*LEARJET*'           { 'LEARJET INC.' }
        '*PLACETECO*'         { 'PLACETECO INC.' }
        '*SATAIR*'            { 'SATAIR USA INC.' }
        '*CVTCORP*'           { 'CVTCORP' }
        '*FIGEAC*'            { 'FIGEAC AERO France' }
        '*HUTCHISON*'         { 'Hutchison' }
        '*APEX*'              { 'Apex Précision' }
        '*DYNOMAX*'           { 'DYNOMAX INC.' }
        '*TRIUMPH*'           { 'TRIUMPH AEROSTRUCTURES LLC' }
        '*QUALUM*'            { 'QUALUM' }
        '*PRESTIGE*'          { 'PRESTIGE INC.' }
        default               { $rawClient }
    }

    $qty = if ($row['col_qty'] -ne [DBNull]::Value) { [int]$row['col_qty'] } else { 1 }
    $wo = if ($row['col_workOrder'] -ne [DBNull]::Value) { [string]$row['col_workOrder'] } else { '' }
    $comment = if ($row['col_comment'] -ne [DBNull]::Value) { [string]$row['col_comment'] } else { '' }
    $suivi = if ($row['col_suivi'] -ne [DBNull]::Value) { [string]$row['col_suivi'] } else { '' }
    $buyer = if ($row['col_buyer'] -ne [DBNull]::Value) { [string]$row['col_buyer'] } else { '' }

    # Append SUIVI note to comment if present
    $fullComment = $comment
    if ($suivi -and $suivi.Trim()) {
        if ($fullComment) { $fullComment += "`n" }
        $fullComment += "SUIVI: $suivi"
    }

    $orders += @{
        po           = $po.Trim()
        item         = $item.Trim()
        commandeItem = $commandeItem
        client       = $client
        partNumber   = $partNumber.Trim()
        dateRequise  = $dateRequise
        qty          = $qty
        workOrder    = $wo.Trim()
        comment      = $fullComment.Trim()
        buyer        = $buyer.Trim()
    }
}

Write-Host "Mapped $($orders.Count) unique orders" -ForegroundColor Green

# ── Step 5: Get existing orders from Azure ──
Write-Host "Fetching existing orders from Azure..." -ForegroundColor Yellow

try {
    $existing = Invoke-RestMethod -Uri "$API_BASE/order-list" -Method Get
    $existingKeys = @{}
    foreach ($e in $existing) {
        $existingKeys[$e.commandeItem] = $true
    }
    Write-Host "Found $($existing.Count) existing orders in Azure" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not fetch existing orders: $($_.Exception.Message)" -ForegroundColor Red
    $existingKeys = @{}
}

# ── Step 6: Push to Azure API ──
$newCount = 0
$updateCount = 0
$errorCount = 0
$total = $orders.Count
$i = 0

Write-Host "Syncing orders to Azure..." -ForegroundColor Yellow

foreach ($order in $orders) {
    $i++
    $isNew = -not $existingKeys.ContainsKey($order.commandeItem)
    $action = if ($isNew) { 'NEW' } else { 'UPD' }

    try {
        $json = $order | ConvertTo-Json -Compress
        $response = Invoke-RestMethod -Uri "$API_BASE/order-save" `
            -Method Post `
            -ContentType 'application/json' `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($json))

        if ($isNew) { $newCount++ } else { $updateCount++ }

        # Progress every 50 orders
        if ($i % 50 -eq 0 -or $i -eq $total) {
            $pct = [math]::Round(($i / $total) * 100)
            Write-Host "  [$pct%] $i / $total  (new: $newCount, updated: $updateCount)" -ForegroundColor Gray
        }
    } catch {
        $errorCount++
        Write-Host "  ERROR [$action] $($order.commandeItem): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ── Step 7: Summary ──
Write-Host "`n=== Sync Complete ===" -ForegroundColor Cyan
Write-Host "  Total from SQL:  $($orders.Count)" -ForegroundColor White
Write-Host "  New orders:      $newCount" -ForegroundColor Green
Write-Host "  Updated orders:  $updateCount" -ForegroundColor Yellow
Write-Host "  Errors:          $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })
Write-Host ""
