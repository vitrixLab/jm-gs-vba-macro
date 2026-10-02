# test_pj_cdj_sync.ps1
$ErrorActionPreference = "Stop"
$xlsmPath = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"

Write-Host "Opening workbook..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $true

try {
    $wb = $xl.Workbooks.Open($xlsmPath, 0, $false)
    $xl.EnableEvents = $true

    Write-Host "Calling RefreshGL..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $xl.Run("RefreshGL")
    $sw.Stop()
    Write-Host "RefreshGL completed in $($sw.Elapsed.TotalSeconds)s"

    $wsAudit = $wb.Sheets.Item("GL_AUDIT")
    Write-Host "GL_AUDIT Status: $($wsAudit.Cells.Item(7, 3).Value), Diff: $($wsAudit.Cells.Item(4, 3).Value), Unmapped: $($wsAudit.Cells.Item(5, 3).Value)"

    Write-Host "`nTesting FindExpenseColumn..."
    $testMap = @(
        @{ Coa = "Supplies"; Expected = 14 },
        @{ Coa = "Clinic Materials and Supplies"; Expected = 14 },
        @{ Coa = "Rental"; Expected = 15 },
        @{ Coa = "Rent Expense"; Expected = 15 },
        @{ Coa = "Fuel and Oil"; Expected = 16 },
        @{ Coa = "Communication"; Expected = 17 },
        @{ Coa = "Light and Water Expense"; Expected = 17 },
        @{ Coa = "Professional Fees"; Expected = 18 },
        @{ Coa = "Clinicians Fee - Visiting"; Expected = 18 },
        @{ Coa = "Miscellaneous"; Expected = 19 },
        @{ Coa = "Representation"; Expected = 19 },
        @{ Coa = "Pantry Supplies"; Expected = 19 }
    )
    foreach ($t in $testMap) {
        $col = $xl.Run("FindExpenseColumn", $t.Coa)
        $status = if ($col -eq $t.Expected) { "OK" } else { "MISMATCH (got $col, expected " + $t.Expected + ")" }
        Write-Host "  COA '$($t.Coa)' -> Col $col [$status]"
    }

    Write-Host "`nTesting PJ -> CDJ Sync..."
    $wsPJ = $wb.Sheets.Item("PJ")
    $wsCDJ = $wb.Sheets.Item("CDJ")

    # Find next available row in PJ
    $pjRow = 10
    while ($null -ne $wsPJ.Cells.Item($pjRow, 4).Value -and "" -ne $wsPJ.Cells.Item($pjRow, 4).Value) {
        $pjRow++
    }
    Write-Host "  Using PJ row $pjRow for live sync test..."

    $refId = "PJ-SYNC-TEST"
    $wsPJ.Cells.Item($pjRow, 2).Value = "JAN"
    $wsPJ.Cells.Item($pjRow, 4).Value = "233-251-708"
    $wsPJ.Cells.Item($pjRow, 5).Value = "FEDERAL BRENT RETAIL INC"
    $wsPJ.Cells.Item($pjRow, 7).Value = "Clinic Materials and Supplies"
    $wsPJ.Cells.Item($pjRow, 8).Value = $refId
    $wsPJ.Cells.Item($pjRow, 10).Value = 5000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 600.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 5600.00

    Write-Host "  Calling SyncPJToCDJ..."
    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cdjTargetRow = $xl.Run("FindCDJRow", $refId)
    Write-Host "  CDJ target row: $cdjTargetRow"

    if ($cdjTargetRow -lt 15) { throw "Sync failed - target row not found" }

    $cdjPart = $wsCDJ.Cells.Item($cdjTargetRow, 5).Value
    $cdjCash = $wsCDJ.Cells.Item($cdjTargetRow, 6).Value
    $cdjVAT  = $wsCDJ.Cells.Item($cdjTargetRow, 7).Value
    $cdjMat  = $wsCDJ.Cells.Item($cdjTargetRow, 14).Value
    $cdjDeb  = $wsCDJ.Cells.Item($cdjTargetRow, 20).Value
    Write-Host "  CDJ Row $($cdjTargetRow): Part='$cdjPart', Cash=$cdjCash, VAT=$cdjVAT, ClinicMat=$cdjMat, Debit=$cdjDeb"

    if ($cdjPart -ne "FEDERAL BRENT RETAIL INC") { throw "Particulars mismatch" }
    if ([Math]::Abs($cdjCash - (-5600.00)) -gt 0.01) { throw "Cash mismatch" }
    if ([Math]::Abs($cdjVAT - 600.00) -gt 0.01) { throw "VAT mismatch" }
    if ([Math]::Abs($cdjMat - 5000.00) -gt 0.01) { throw "Clinic Materials mismatch" }
    if ([Math]::Abs($cdjDeb - 5600.00) -gt 0.01) { throw "Debit mismatch" }
    Write-Host "  [PASS] PJ -> CDJ INSERT test SUCCESSFUL!"

    # Test Update
    Write-Host "`nTesting PJ -> CDJ UPDATE..."
    $wsPJ.Cells.Item($pjRow, 10).Value = 10000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 1200.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 11200.00

    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cdjCashUpd = $wsCDJ.Cells.Item($cdjTargetRow, 6).Value
    $cdjMatUpd  = $wsCDJ.Cells.Item($cdjTargetRow, 14).Value
    $cdjDebUpd  = $wsCDJ.Cells.Item($cdjTargetRow, 20).Value
    Write-Host "  CDJ Row $cdjTargetRow after update: Cash=$cdjCashUpd, ClinicMat=$cdjMatUpd, Debit=$cdjDebUpd"

    if ([Math]::Abs($cdjCashUpd - (-11200.00)) -gt 0.01) { throw "Updated Cash mismatch" }
    if ([Math]::Abs($cdjMatUpd - 10000.00) -gt 0.01) { throw "Updated ClinicMat mismatch" }
    if ([Math]::Abs($cdjDebUpd - 11200.00) -gt 0.01) { throw "Updated Debit mismatch" }
    Write-Host "  [PASS] PJ -> CDJ UPDATE test SUCCESSFUL!"

    # Clean up test rows
    Write-Host "`nCleaning up test rows..."
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjTargetRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value = $cdjTargetRow
    $wb.Save()
    Write-Host "  Cleanup complete and workbook saved."

    # Final RefreshGL check
    $xl.Run("RefreshGL")
    Write-Host "`nFinal GL_AUDIT Status: $($wsAudit.Cells.Item(7, 3).Value)"
    Write-Host "=== ALL PJ -> CDJ SYNC TESTS PASSED CERTIFIED! ==="

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
