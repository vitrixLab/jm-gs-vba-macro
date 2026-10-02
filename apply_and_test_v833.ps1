# apply_and_test_v833.ps1
# Applies modEngine_v833_hardened.bas, Sheet12_v833_hardened.cls, Sheet11_v833_hardened.cls, modGLAggregation.bas
# and runs end-to-end PJ->CDJ sync and GL refresh tests.

$ErrorActionPreference = "Stop"
$xlsmPath = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"
$modEngine = "D:\citrixlabph\globalsmile\modEngine_v833_hardened.bas"
$sheet12 = "D:\citrixlabph\globalsmile\Sheet12_v833_hardened.cls"
$sheet11 = "D:\citrixlabph\globalsmile\Sheet11_v833_hardened.cls"
$modGLAgg = "D:\citrixlabph\globalsmile\modGLAggregation.bas"

Write-Host "=== Starting Excel Automation ==="
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false

try {
    Write-Host "Opening $xlsmPath..."
    $wb = $xl.Workbooks.Open($xlsmPath, 0, $false)
    $vbp = $wb.VBProject

    # 1. Update modEngine
    Write-Host "Updating modEngine..."
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Name -eq "modEngine") {
            $vbp.VBComponents.Remove($comp)
            Write-Host "  Removed old modEngine."
            break
        }
    }
    $vbp.VBComponents.Import($modEngine) | Out-Null
    Write-Host "  Imported modEngine_v833_hardened.bas."

    # 2. Update modGLAggregation
    Write-Host "Updating modGLAggregation..."
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Name -eq "modGLAggregation") {
            $vbp.VBComponents.Remove($comp)
            Write-Host "  Removed old modGLAggregation."
            break
        }
    }
    if (Test-Path $modGLAgg) {
        $vbp.VBComponents.Import($modGLAgg) | Out-Null
        Write-Host "  Imported modGLAggregation.bas."
    }

    # 3. Update Sheet12 (PJ tab)
    Write-Host "Updating Sheet12 (PJ tab)..."
    $s12Raw = Get-Content $sheet12 -Raw
    $s12CleanLines = @()
    foreach ($line in ($s12Raw -split "`r?`n")) {
        if ($line -notmatch "^Attribute\s+") {
            $s12CleanLines += $line
        }
    }
    $s12CleanCode = $s12CleanLines -join "`r`n"
    $comp12 = $vbp.VBComponents.Item("Sheet12")
    $comp12.CodeModule.DeleteLines(1, $comp12.CodeModule.CountOfLines)
    $comp12.CodeModule.AddFromString($s12CleanCode)
    Write-Host "  Updated Sheet12 CodeModule."

    # 4. Update Sheet11 (PJ Non-Vat tab)
    Write-Host "Updating Sheet11 (PJ Non-Vat tab)..."
    $s11Raw = Get-Content $sheet11 -Raw
    $s11CleanLines = @()
    foreach ($line in ($s11Raw -split "`r?`n")) {
        if ($line -notmatch "^Attribute\s+") {
            $s11CleanLines += $line
        }
    }
    $s11CleanCode = $s11CleanLines -join "`r`n"
    $comp11 = $vbp.VBComponents.Item("Sheet11")
    $comp11.CodeModule.DeleteLines(1, $comp11.CodeModule.CountOfLines)
    $comp11.CodeModule.AddFromString($s11CleanCode)
    Write-Host "  Updated Sheet11 CodeModule."

    # Save
    $wb.Save()
    Write-Host "Workbook saved with all updated components."

    # Enable events for testing
    $xl.EnableEvents = $true

    # 5. Test InitializeEngine
    Write-Host "`n--- Testing modEngine.InitializeEngine ---"
    $xl.Run("InitializeEngine", "CDJ")
    Write-Host "  InitializeEngine('CDJ') executed successfully."

    # 6. Test FindExpenseColumn
    Write-Host "`n--- Testing modEngine.FindExpenseColumn ---"
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
        $status = if ($col -eq $t.Expected) { "OK" } else { "MISMATCH (expected " + $t.Expected + ")" }
        Write-Host "  COA '$($t.Coa)' -> Col $col [$status]"
    }

    # 7. Test PJ -> CDJ Sync
    Write-Host "`n--- Testing PJ -> CDJ Sync Live ---"
    $wsPJ = $wb.Sheets.Item("PJ")
    $wsCDJ = $wb.Sheets.Item("CDJ")

    $pjRow = 10
    while ($null -ne $wsPJ.Cells.Item($pjRow, 4).Value -and "" -ne $wsPJ.Cells.Item($pjRow, 4).Value) {
        $pjRow++
    }
    Write-Host "  Using PJ row $pjRow for live test..."

    $refId = "PJ-TEST999"
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
    Write-Host "  CDJ target row for $($refId): $cdjTargetRow"
    if ($cdjTargetRow -lt 15) { throw "Target row not found in CDJ" }

    $cdjPart = $wsCDJ.Cells.Item($cdjTargetRow, 5).Value
    $cdjCash = $wsCDJ.Cells.Item($cdjTargetRow, 6).Value
    $cdjVAT  = $wsCDJ.Cells.Item($cdjTargetRow, 7).Value
    $cdjMat  = $wsCDJ.Cells.Item($cdjTargetRow, 14).Value
    $cdjDeb  = $wsCDJ.Cells.Item($cdjTargetRow, 20).Value
    Write-Host "  CDJ Row $cdjTargetRow : Part='$cdjPart', Cash=$cdjCash, VAT=$cdjVAT, Mat=$cdjMat, Debit=$cdjDeb"

    if ($cdjPart -ne "FEDERAL BRENT RETAIL INC") { throw "Particulars mismatch: $cdjPart" }
    if ([Math]::Abs($cdjCash - (-5600.00)) -gt 0.01) { throw "Cash mismatch: $cdjCash" }
    if ([Math]::Abs($cdjVAT - 600.00) -gt 0.01) { throw "VAT mismatch: $cdjVAT" }
    if ([Math]::Abs($cdjMat - 5000.00) -gt 0.01) { throw "Clinic Materials mismatch: $cdjMat" }
    if ([Math]::Abs($cdjDeb - 5600.00) -gt 0.01) { throw "Debit mismatch: $cdjDeb" }
    Write-Host "  [PASS] PJ -> CDJ INSERT test successful!"

    # Test Update
    Write-Host "`n--- Testing PJ -> CDJ UPDATE ---"
    $wsPJ.Cells.Item($pjRow, 10).Value = 10000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 1200.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 11200.00

    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cdjCashUpd = $wsCDJ.Cells.Item($cdjTargetRow, 6).Value
    $cdjMatUpd  = $wsCDJ.Cells.Item($cdjTargetRow, 14).Value
    $cdjDebUpd  = $wsCDJ.Cells.Item($cdjTargetRow, 20).Value
    Write-Host "  CDJ Row $cdjTargetRow after update: Cash=$cdjCashUpd, Mat=$cdjMatUpd, Debit=$cdjDebUpd"

    if ([Math]::Abs($cdjCashUpd - (-11200.00)) -gt 0.01) { throw "Updated Cash mismatch" }
    if ([Math]::Abs($cdjMatUpd - 10000.00) -gt 0.01) { throw "Updated Mat mismatch" }
    if ([Math]::Abs($cdjDebUpd - 11200.00) -gt 0.01) { throw "Updated Debit mismatch" }
    Write-Host "  [PASS] PJ -> CDJ UPDATE test successful!"

    # Clean up
    Write-Host "`n--- Cleaning up test rows ---"
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjTargetRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value = $cdjTargetRow
    $wb.Save()
    Write-Host "  Cleaned up."

    # 8. Test RefreshGL
    Write-Host "`n--- Testing RefreshGL ---"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $xl.Run("RefreshGL")
    $sw.Stop()
    Write-Host "  RefreshGL completed in $($sw.Elapsed.TotalSeconds)s"

    $wsAudit = $wb.Sheets.Item("GL_AUDIT")
    $status = $wsAudit.Cells.Item(7, 3).Value
    $diff   = $wsAudit.Cells.Item(4, 3).Value
    $unmap  = $wsAudit.Cells.Item(5, 3).Value
    Write-Host "  GL_AUDIT Status: $status, Difference: $diff, Unmapped: $unmap"

    $wsGL = $wb.Sheets.Item("GL")
    $cibCr = $wsGL.Cells.Item(13, 10).Value
    $miscDb = $wsGL.Cells.Item(520, 9).Value
    Write-Host "  GL Cash In Bank Jan Credit: $cibCr"
    Write-Host "  GL Miscellaneous Jan Debit: $miscDb"

    Write-Host "`n=== ALL TESTS PASSED! ==="

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
