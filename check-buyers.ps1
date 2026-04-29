$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = 'Server=WIN2K19\SQLEXPRESS;Database=DMG_BASE_TAKe2;Trusted_Connection=Yes;'
$conn.Open()

# Get column names
$livCmd = $conn.CreateCommand()
$livCmd.CommandText = 'SELECT TOP 1 * FROM [*REG_LIVRAISON]'
$livA = New-Object System.Data.SqlClient.SqlDataAdapter($livCmd)
$livDt = New-Object System.Data.DataTable
$livA.Fill($livDt) | Out-Null
$numCol = $livDt.Columns[1].ColumnName
$dateReqCol = $livDt.Columns[5].ColumnName
$cancelCol = $livDt.Columns[14].ColumnName

$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT a.NOM_ACHETEUR, COUNT(*) as cnt FROM [*REG_PO] p JOIN [*REG_ITEM] i ON i.po_ident = p.po_ident JOIN [*REG_LIVRAISON] l ON l.[$numCol] = i.[$numCol] LEFT JOIN ACHETEURS a ON a.ACHETEUR_ID = p.ACHETEUR_ID WHERE ISNULL(l.[$cancelCol],0)=0 AND ISNULL(l.ARCHIVE,0)=0 AND l.[$dateReqCol] >= '2025-01-01' GROUP BY a.NOM_ACHETEUR ORDER BY cnt DESC"
$a = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
$a.Fill($dt) | Out-Null
Write-Host "Buyer distribution (active orders >= 2025):"
foreach ($row in $dt.Rows) {
    Write-Host "  $($row['NOM_ACHETEUR']) = $($row['cnt'])"
}

$conn.Close()
