# ══════════════════════════════════════════════════════════════
# DMG Follow-Up - Send Carco Reports via Outlook
# Fetches orders from Azure API, generates Excel per buyer, emails via Outlook
# Usage: powershell -ExecutionPolicy Bypass -File send-carco.ps1
# ══════════════════════════════════════════════════════════════

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$API_BASE = 'https://dmg-followup-proxy.azurewebsites.net/api/followup'
$AUTH_BASE = 'https://dmg-proxy-b4fpgjh8begye3fb.canadacentral-01.azurewebsites.net/api/dmg'

# ── Auth (set DMG_USER + DMG_PASS env vars for unattended runs) ──
$DMG_USER = if ($env:DMG_USER) { $env:DMG_USER } else { Read-Host "DMG username" }
$DMG_PASS = if ($env:DMG_PASS) {
    $env:DMG_PASS
} else {
    $sec = Read-Host "DMG password" -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
}
$loginBody = @{ username = $DMG_USER; password = $DMG_PASS; site = 'DMG' } | ConvertTo-Json -Compress
try {
    $loginRes = Invoke-RestMethod -Uri "$AUTH_BASE/auth-login" -Method Post -ContentType 'application/json' -Body $loginBody
    $TOKEN = $loginRes.token
} catch { Write-Host "Login failed: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }
$AUTH_HEADERS = @{ Authorization = "Bearer $TOKEN" }

# ── BUYER EMAILS - Fill in the email addresses ──
$buyerEmails = @{
    # ── TEST MODE: all emails go to JP ──
    'ANIK LAURIN'           = ''  # PROD: anik.laurin@airbus.com
    'LAURA DASCALESCU'      = 'jpdor9@gmail.com'  # PROD: Laura.Dascalescu@aero.bombardier.com
    'CHAYMAE A. BENNANI'    = ''  # PROD: chaymae.abbanabennani@aero.bombardier.com
    'CAROLINE MASSIA'       = ''  # PROD: caroline-diane-eva-rachel.massia@airbus.com
    'CHANTAL PAQUET'        = ''  # PROD: cpaquet@placeteco.com
    'FRANCIS VARELA'        = ''  # PROD: francis.varela@aero.bombardier.com
    'ROSTOM MAAROUFI'       = ''  # PROD: rostom.maaroufi@airbus.com
    'VERED HABA'            = ''  # PROD: Vered.Haba@aero.bombardier.com
    'NOUR MALKI'            = ''  # PROD: nour.malki@aero.bombardier.com
    'JADEN HARP'            = ''  # PROD: jaden.harp@aero.bombardier.com
    'ERIC BERGOIN'          = ''  # PROD: eric.bergoin@aero.bombardier.com
    'TRACI LIN EDWARDS'     = ''  # PROD: tred@satair.com
    'MORAYMA URRUTIA'       = ''  # PROD: morayma.urrutia@airbus.com
    'ERIC THEBERGE'         = ''  # PROD: eric.theberge@aero.bombardier.com
    'KERRI GORDON'          = ''  # PROD: kgo@satair.com
    'SERGIO VALDEZ'         = ''  # PROD: rafael.sergio.valdez@aero.bombardier.com
    'ALEXANDER GAUDIO'      = ''  # PROD: bba.aogparts@aero.bombardier.com
    'CAROLINE GRAVEL'       = ''  # PROD: cgravel@cvtcorp.com
    'GABRIELA TALAVERA'     = ''
}

# CC yourself on every email (leave empty to skip)
$ccEmail = ''

# ── Step 1: Fetch orders from Azure API ──
Write-Host "`n=== DMG Carco Email Sender ===" -ForegroundColor Cyan
Write-Host "Fetching orders from Azure..." -ForegroundColor Yellow

$orders = Invoke-RestMethod -Uri "$API_BASE/order-list" -Method Get -Headers $AUTH_HEADERS
$recipes = Invoke-RestMethod -Uri "$API_BASE/recipe-list" -Method Get -Headers $AUTH_HEADERS
$suiviRaw = Invoke-RestMethod -Uri "$API_BASE/suivi-list" -Method Get -Headers $AUTH_HEADERS

# Build suivi lookup
$suivi = @{}
foreach ($s in $suiviRaw) {
    if ($s.commandeItem) { $suivi[$s.commandeItem] = $s }
}

Write-Host "Loaded $($orders.Count) orders, $($recipes.Count) recipes" -ForegroundColor Green

# ── Step 2: Filter active orders (same logic as Carco UI) ──
$today = (Get-Date).Date
$pastCutoff = $today.AddDays(-30)

$activeOrders = $orders | Where-Object {
    -not $_.archived -and
    $_.partNumber -and
    $_.partNumber.ToUpper() -ne 'FAI' -and
    $_.partNumber.ToUpper() -ne 'NRC' -and
    $_.partNumber.ToUpper() -ne 'NCR' -and
    -not $_.partNumber.ToUpper().StartsWith('EXPEDITE') -and
    $_.dateRequise -and
    ([DateTime]::Parse($_.dateRequise) -ge $pastCutoff) -and
    $_.buyer
}

Write-Host "Active orders for Carco: $($activeOrders.Count)" -ForegroundColor Green

# ── Step 3: Group by buyer ──
$byBuyer = $activeOrders | Group-Object -Property buyer

# ── Step 4: Generate Excel and send per buyer ──
Write-Host "Starting Outlook..." -ForegroundColor Yellow
$outlook = New-Object -ComObject Outlook.Application

$exportDir = Join-Path $PSScriptRoot 'carco-exports'
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

$sentCount = 0
$skippedCount = 0

foreach ($group in $byBuyer) {
    $buyerName = $group.Name
    $email = $buyerEmails[$buyerName]

    if (-not $email) {
        Write-Host "  SKIP $buyerName - no email configured" -ForegroundColor DarkGray
        $skippedCount++
        continue
    }

    $buyerOrders = $group.Group | Sort-Object { $_.dateRequise }

    # Build rows
    $rows = @()
    foreach ($o in $buyerOrders) {
        $recipe = $recipes | Where-Object { $_.partNumber -eq $o.partNumber } | Select-Object -First 1
        $nbOps = 0
        $done = 0
        if ($recipe -and $recipe.ops) {
            $nbOps = $recipe.ops.Count
            $statuses = $suivi[$o.commandeItem]
            for ($i = 0; $i -lt $nbOps; $i++) {
                $st = ''
                if ($statuses) {
                    $key = "op$($i+1)"
                    $st = $statuses.$key
                }
                if ($st -eq 'Complété' -or $st -eq 'N/A') { $done++ }
            }
        }
        $pct = if ($nbOps -gt 0) { [math]::Round(($done / $nbOps) * 100) } else { 0 }

        # Status logic
        $dr = [DateTime]::Parse($o.dateRequise)
        $daysLeft = ($dr - $today).Days
        if ($o.carcoStatus) {
            $status = $o.carcoStatus
        } elseif ($daysLeft -lt 0) {
            $status = 'Late'
        } elseif ($daysLeft -le 7) {
            $status = 'At risk'
        } else {
            $status = 'On schedule'
        }

        $rows += [PSCustomObject]@{
            'Commande'    = $o.commandeItem
            'No Pièce'    = $o.partNumber
            'Client'      = $o.client
            'Date Client' = $o.dateRequise
            'Progression' = "$pct%"
            'Status'      = $status
            'New ECD'     = if ($status -eq 'Late') { $o.carcoEcd } else { '' }
            'Note'        = $o.carcoNote
        }
    }

    # Export to Excel (CSV as fallback-safe, but we use xlsx via COM)
    $dateStr = $today.ToString('yyyy-MM-dd')
    $safeName = $buyerName -replace '[^a-zA-Z0-9]', '_'
    $filename = "Carco_${safeName}_${dateStr}.xlsx"
    $filepath = Join-Path $exportDir $filename

    # Use Excel COM to create proper xlsx
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Carco'

        # Title row
        $ws.Cells.Item(1, 1).Value2 = "DMG Aerospace - Rapport Carco - $buyerName - $dateStr"
        $ws.Range("A1:H1").MergeCells = $true
        $ws.Cells.Item(1, 1).Font.Size = 14
        $ws.Cells.Item(1, 1).Font.Bold = $true

        # Headers
        $headers = @('Commande', 'No Pièce', 'Client', 'Date Client', 'Progression', 'Status', 'New ECD', 'Note')
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $ws.Cells.Item(3, $c + 1).Value2 = $headers[$c]
            $ws.Cells.Item(3, $c + 1).Font.Bold = $true
            $ws.Cells.Item(3, $c + 1).Interior.ColorIndex = 48
            $ws.Cells.Item(3, $c + 1).Font.ColorIndex = 2
        }

        # Data rows
        $r = 4
        foreach ($row in $rows) {
            $ws.Cells.Item($r, 1).Value2 = $row.Commande
            $ws.Cells.Item($r, 2).Value2 = $row.'No Pièce'
            $ws.Cells.Item($r, 3).Value2 = $row.Client
            $ws.Cells.Item($r, 4).Value2 = $row.'Date Client'
            $ws.Cells.Item($r, 5).Value2 = $row.Progression
            $ws.Cells.Item($r, 6).Value2 = $row.Status

            # Color status cell (RGB as OLE: B*65536 + G*256 + R)
            if ($row.Status -eq 'Late') {
                $ws.Cells.Item($r, 6).Font.Color = 255        # Red
            } elseif ($row.Status -eq 'At risk') {
                $ws.Cells.Item($r, 6).Font.Color = 33023      # Orange
            } else {
                $ws.Cells.Item($r, 6).Font.Color = 32768      # Green
            }

            $ws.Cells.Item($r, 7).Value2 = $row.'New ECD'
            $ws.Cells.Item($r, 8).Value2 = $row.Note
            $r++
        }

        # Auto-fit columns
        $ws.Columns.Item("A:H").AutoFit() | Out-Null

        $wb.SaveAs($filepath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Host "  ERROR creating Excel for $buyerName : $($_.Exception.Message)" -ForegroundColor Red
        try { $excel.Quit() } catch {}
        continue
    }

    # Send email via Outlook
    try {
        $mail = $outlook.CreateItem(0)
        $mail.To = $email
        if ($ccEmail) { $mail.CC = $ccEmail }
        $mail.Subject = "DMG Aerospace - Rapport Carco - $buyerName - $dateStr"
        $mail.Body = @"
Hello,

Please find attached the advancement report for your current orders at DMG Aerospace.

Orders in this report: $($rows.Count)

Should you have any questions, please do not hesitate to contact us.

Best regards,
DMG Aerospace - Usinage mécanique DMG inc.
"@
        $mail.Attachments.Add($filepath) | Out-Null
        $mail.Send()

        Write-Host "  SENT $buyerName ($email) - $($rows.Count) orders" -ForegroundColor Green
        $sentCount++
    } catch {
        Write-Host "  ERROR sending to $buyerName : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ── Step 5: Summary ──
Write-Host "`n=== Carco Emails Complete ===" -ForegroundColor Cyan
Write-Host "  Sent:    $sentCount" -ForegroundColor Green
Write-Host "  Skipped: $skippedCount (no email)" -ForegroundColor Yellow
Write-Host "  Files:   $exportDir" -ForegroundColor Gray
Write-Host ""
