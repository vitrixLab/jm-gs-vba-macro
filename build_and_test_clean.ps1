# build_and_test_clean.ps1
# Clean in-place code update and full certification of v8.3.3 from v8.3-FINAL base

$ErrorActionPreference = "Stop"
$log = "D:\citrixlabph\globalsmile\build_clean.log"
function L($m) { $ts = Get-Date -f "HH:mm:ss"; "$ts  $m" | Tee-Object -Append -FilePath $log; Write-Host "$ts  $m" }

"" | Out-File $log
L "================================================================="
L "  BUILD & TEST CLEAN: Global-Smile_2026-v8.3.3.xlsm"
L "================================================================="

$baseXlsm    = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3-FINAL.xlsm"
$targetXlsm  = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"

$modEngineSrc = "D:\citrixlabph\globalsmile\modEngine_v833_hardened.bas"
$sheet12Src   = "D:\citrixlabph\globalsmile\Sheet12_v833_hardened.cls"
$sheet11Src   = "D:\citrixlabph\globalsmile\Sheet11_v833_hardened.cls"

$modGLMap     = "D:\citrixlabph\globalsmile\modGLWorkbookMap.bas"
$modGLScan    = "D:\citrixlabph\globalsmile\modGLScanEngine.bas"
$modGLRef     = "D:\citrixlabph\globalsmile\modGLRefresh.bas"
$modGLGate    = "D:\citrixlabph\globalsmile\modGLGate.bas"
$modGLAgg     = "D:\citrixlabph\globalsmile\modGLAggregation.bas"
$modPJSync    = "D:\citrixlabph\globalsmile\modPJSync.bas"

# Step 1: Copy base
L "Step 1: Copying base workbook..."
Copy-Item -Path $baseXlsm -Destination $targetXlsm -Force
L "  Copied."

# Step 2: Open Excel COM
L "Step 2: Launching Excel COM..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible           = $false
$xl.DisplayAlerts     = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents      = $false

try {
    L "Step 3: Opening target workbook..."
    $wb = $xl.Workbooks.Open($targetXlsm, 0, $false)
    $vbp = $wb.VBProject

    # Clean Sheet9 backticks
    L "Step 4: Cleaning Sheet9..."
    $comp9 = $vbp.VBComponents.Item("Sheet9")
    $s9Code = $comp9.CodeModule.Lines(1, $comp9.CodeModule.CountOfLines)
    if ($s9Code -match '```') {
        $cleanedS9 = ($s9Code -split "`r?`n" | Where-Object { $_ -notmatch '^```' }) -join "`r`n"
        $comp9.CodeModule.DeleteLines(1, $comp9.CodeModule.CountOfLines)
        $comp9.CodeModule.AddFromString($cleanedS9)
        L "  Sheet9 cleaned."
    }

    # Update Sheet12 (PJ tab)
    L "Step 5: Updating Sheet12 code..."
    $s12Raw = Get-Content $sheet12Src -Raw
    $s12Clean = (($s12Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp12 = $vbp.VBComponents.Item("Sheet12")
    $comp12.CodeModule.DeleteLines(1, $comp12.CodeModule.CountOfLines)
    $comp12.CodeModule.AddFromString($s12Clean)
    L "  Sheet12 updated."

    # Update Sheet11 (PJ Non-Vat tab)
    L "Step 6: Updating Sheet11 code..."
    $s11Raw = Get-Content $sheet11Src -Raw
    $s11Clean = (($s11Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp11 = $vbp.VBComponents.Item("Sheet11")
    $comp11.CodeModule.DeleteLines(1, $comp11.CodeModule.CountOfLines)
    $comp11.CodeModule.AddFromString($s11Clean)
    L "  Sheet11 updated."

    # Update modEngine in place
    L "Step 7: Updating modEngine code in place..."
    $meRaw = Get-Content $modEngineSrc -Raw
    $meClean = (($meRaw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $compME = $vbp.VBComponents.Item("modEngine")
    $compME.CodeModule.DeleteLines(1, $compME.CodeModule.CountOfLines)
    $compME.CodeModule.AddFromString($meClean)
    L "  modEngine updated."

    # Remove conflicting old modules
    L "Step 8: Removing old legacy modules..."
    $toRemove = @("modGLRefresh", "Module1", "modPJAutomation")
    foreach ($name in $toRemove) {
        try {
            $c = $vbp.VBComponents.Item($name)
            if ($null -ne $c) {
                $vbp.VBComponents.Remove($c)
                L "  Removed $name"
            }
        } catch {}
    }

    # Import new hardened modules
    L "Step 9: Importing new GL and Sync modules..."
    $mods = @($modGLMap, $modGLScan, $modGLRef, $modGLGate, $modGLAgg, $modPJSync)
    foreach ($m in $mods) {
        if (Test-Path $m) {
            $vbp.VBComponents.Import($m) | Out-Null
            L "  Imported $([System.IO.Path]::GetFileName($m))"
        }
    }

    # Enable events
    $xl.EnableEvents = $true

    # Step 10: Run RefreshGL()
    L "Step 10: Running RefreshGL()..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $xl.Run("RefreshGL")
    $sw.Stop()
    L "  RefreshGL() finished in $($sw.Elapsed.TotalSeconds)s"

    $wsAudit = $wb.Sheets.Item("GL_AUDIT")
    $status = $wsAudit.Cells.Item(7, 3).Value
    $diff   = $wsAudit.Cells.Item(4, 3).Value
    $unmap  = $wsAudit.Cells.Item(5, 3).Value
    $inv    = $wsAudit.Cells.Item(6, 3).Value
    L "  GL_AUDIT Status: $status | Diff: $diff | Unmapped: $unmap | Invalid: $inv"

    if ($status -ne "PASS") { throw "GL_AUDIT Status is not PASS: $status" }

    # Step 11: Test PJ -> CDJ Live Sync
    L "Step 11: Testing PJ -> CDJ Live Sync..."
    $wsPJ = $wb.Sheets.Item("PJ")
    $wsCDJ = $wb.Sheets.Item("CDJ")

    $pjRow = 10
    while ($null -ne $wsPJ.Cells.Item($pjRow, 4).Value -and "" -ne $wsPJ.Cells.Item($pjRow, 4).Value) {
        $pjRow++
    }
    L "  Using PJ Row $pjRow for test..."

    $refId = "PJ-TEST-833"
    $wsPJ.Cells.Item($pjRow, 2).Value = "JAN"
    $wsPJ.Cells.Item($pjRow, 4).Value = "233-251-708"
    $wsPJ.Cells.Item($pjRow, 5).Value = "FEDERAL BRENT RETAIL INC"
    $wsPJ.Cells.Item($pjRow, 7).Value = "Clinic Materials and Supplies"
    $wsPJ.Cells.Item($pjRow, 8).Value = $refId
    $wsPJ.Cells.Item($pjRow, 10).Value = 5000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 600.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 5600.00

    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)
    $cdjRow = $xl.Run("FindCDJRow", $refId)
    L "  CDJ Target Row: $cdjRow"
    if ($cdjRow -lt 15) { throw "Sync INSERT failed: Ref $refId not found in CDJ" }

    $cCash = $wsCDJ.Cells.Item($cdjRow, 6).Value
    $cVAT  = $wsCDJ.Cells.Item($cdjRow, 7).Value
    $cMat  = $wsCDJ.Cells.Item($cdjRow, 14).Value
    $cDeb  = $wsCDJ.Cells.Item($cdjRow, 20).Value
    L "  CDJ Row $($cdjRow): Cash=$cCash, VAT=$cVAT, Mat=$cMat, Debit=$cDeb"

    if ([Math]::Abs($cCash - (-5600.00)) -gt 0.01 -or [Math]::Abs($cMat - 5000.00) -gt 0.01) {
        throw "Sync INSERT value mismatch"
    }
    L "  PASS: PJ -> CDJ INSERT verified."

    # Test update
    $wsPJ.Cells.Item($pjRow, 10).Value = 8000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 960.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 8960.00
    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cCashUpd = $wsCDJ.Cells.Item($cdjRow, 6).Value
    $cMatUpd  = $wsCDJ.Cells.Item($cdjRow, 14).Value
    L "  CDJ Row $cdjRow after update: Cash=$cCashUpd, Mat=$cMatUpd"
    if ([Math]::Abs($cCashUpd - (-8960.00)) -gt 0.01 -or [Math]::Abs($cMatUpd - 8000.00) -gt 0.01) {
        throw "Sync UPDATE value mismatch"
    }
    L "  PASS: PJ -> CDJ UPDATE verified."

    # Cleanup test row
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value = $cdjRow
    L "  Test rows cleaned up."

    # Save finalized workbook
    L "Step 12: Saving finalized certified workbook..."
    $wb.Save()
    L "  Saved."

    $hash = (Get-FileHash -Path $targetXlsm -Algorithm SHA256).Hash
    L "SHA-256: $hash"
    L "================================================================="
    L "  BUILD SUCCESS: ALL PJ->CDJ SYNC & GL REFRESH TESTS PASS 100%"
    L "================================================================="

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
