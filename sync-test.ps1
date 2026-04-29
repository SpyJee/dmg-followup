# Quick dry-run test of SQL sync query
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$DATE_CUTOFF = '2025-01-01'

$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = 'Server=WIN2K19\SQLEXPRESS;Database=DMG_BASE_TAKe2;Trusted_Connection=Yes;'
$conn.Open()
Write-Host "Connected" -ForegroundColor Green

# Get column names dynamically
$t = New-Object System.Data.DataTable
(New-Object System.Data.SqlClient.SqlDataAdapter((& { $c=$conn.CreateCommand(); $c.CommandText='SELECT TOP 1 * FROM [*REG_ITEM]'; $c }))).Fill($t)|Out-Null
$numCol = $t.Columns[1].ColumnName

$t2 = New-Object System.Data.DataTable
(New-Object System.Data.SqlClient.SqlDataAdapter((& { $c=$conn.CreateCommand(); $c.CommandText='SELECT TOP 1 * FROM CLIENTS_FOURNISSEURS'; $c }))).Fill($t2)|Out-Null
$socCol = $t2.Columns[1].ColumnName

$t3 = New-Object System.Data.DataTable
(New-Object System.Data.SqlClient.SqlDataAdapter((& { $c=$conn.CreateCommand(); $c.CommandText='SELECT TOP 1 * FROM [*REG_LIVRAISON]'; $c }))).Fill($t3)|Out-Null
$qtyCol = $t3.Columns[3].ColumnName
$dateReqCol = $t3.Columns[5].ColumnName
$woCol = $t3.Columns[6].ColumnName
$cancelCol = $t3.Columns[14].ColumnName
$suiviCol = $t3.Columns[24].ColumnName

$query = @"
SELECT
    p.[BON DE COMMANDE] AS col_po,
    i.item AS col_item,
    i.[PART NUMBER] AS col_partNumber,
    i.NOM AS col_partName,
    cf.[$socCol] AS col_client,
    l.[$qtyCol] AS col_qty,
    l.[$dateReqCol] AS col_dateRequise,
    l.COMMENTS AS col_comment,
    l.[$suiviCol] AS col_suivi,
    l.[$cancelCol] AS col_cancelled,
    l.ARCHIVE AS col_archived,
    l.[$woCol] AS col_workOrder
FROM [*REG_PO] p
JOIN [*REG_ITEM] i ON i.po_ident = p.po_ident
JOIN [*REG_LIVRAISON] l ON l.[$numCol] = i.[$numCol]
LEFT JOIN CLIENTS_FOURNISSEURS cf ON cf.FOURNISSEUR_ID = p.FOURNISSEUR_ID
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

Write-Host "Total rows from SQL (>= $DATE_CUTOFF): $count" -ForegroundColor Cyan

# Dedup and map
$seen = @{}
$orders = @()
foreach ($row in $dt.Rows) {
    $po = [string]$row['col_po']
    $item = [string]$row['col_item']
    if ([string]::IsNullOrWhiteSpace($po)) { continue }
    $ci = "$po-$item"
    if ($seen.ContainsKey($ci)) { continue }
    $seen[$ci] = $true

    $dateReq = ''
    if ($row['col_dateRequise'] -ne [DBNull]::Value) {
        $dateReq = ([DateTime]$row['col_dateRequise']).ToString('yyyy-MM-dd')
    }
    $client = if ($row['col_client'] -ne [DBNull]::Value) { [string]$row['col_client'] } else { '(none)' }
    $qty = if ($row['col_qty'] -ne [DBNull]::Value) { [int]$row['col_qty'] } else { 0 }
    $wo = if ($row['col_workOrder'] -ne [DBNull]::Value) { [string]$row['col_workOrder'] } else { '' }
    $pn = if ($row['col_partNumber'] -ne [DBNull]::Value) { [string]$row['col_partNumber'] } else { '' }

    $orders += @{ ci=$ci; client=$client; pn=$pn; qty=$qty; dateReq=$dateReq; wo=$wo }
}

Write-Host "Unique orders after dedup: $($orders.Count)" -ForegroundColor Green
Write-Host ""

# Show first 10 and last 5
Write-Host "=== First 10 orders ===" -ForegroundColor Yellow
$orders[0..9] | ForEach-Object { Write-Host "$($_.ci) | $($_.client) | $($_.pn) | qty=$($_.qty) | $($_.dateReq) | wo=$($_.wo)" }

Write-Host ""
Write-Host "=== Last 5 orders ===" -ForegroundColor Yellow
$orders[($orders.Count-5)..($orders.Count-1)] | ForEach-Object { Write-Host "$($_.ci) | $($_.client) | $($_.pn) | qty=$($_.qty) | $($_.dateReq) | wo=$($_.wo)" }

# Client breakdown
Write-Host ""
Write-Host "=== Client breakdown ===" -ForegroundColor Yellow
$clients = @{}
foreach ($o in $orders) {
    $c = $o.client
    if (-not $clients.ContainsKey($c)) { $clients[$c] = 0 }
    $clients[$c]++
}
foreach ($kv in ($clients.GetEnumerator() | Sort-Object -Property Value -Descending)) {
    Write-Host "  $($kv.Value) - $($kv.Key)"
}
