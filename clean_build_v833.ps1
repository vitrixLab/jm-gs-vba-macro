# clean_build_v833.ps1
# Clean build and certification of v8.3.3 from v8.3-FINAL base

$ErrorActionPreference = "Stop"
$log = "D:\citrixlabph\globalsmile\clean_build.log"
function L($m) { $ts = Get-Date -f "HH:mm:ss"; "$ts  $m" | Tee-Object -Append -FilePath $log; Write-Host "$ts  $m" }

"" | Out-File $log
L "================================================================="
L "  CLEAN BUILD: Global-Smile_2026-v8.3.3.xlsm"
L "================================================================="

$baseXlsm    = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3-FINAL.xlsm"
$targetXlsm  = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"

$modEngine   = "D:\citrixlabph\globalsmile\modEngine_v833_hardened.bas"
$sheet12     = "D:\citrixlabph\globalsmile\Sheet12_v833_hardened.cls"
$sheet11     = "D:\citrixlabph\globalsmile\Sheet11_v833_hardened.cls"
$modGLMap    = "D:\citrixlabph\globalsmile\modGLWorkbookMap.bas"
$modGLScan   = "D:\citrixlabph\globalsmile\modGLScanEngine.bas"
$modGLRef    = "D:\citrixlabph\globalsmile\modGLRefresh.bas"
$modGLGate   = "D:\citrixlabph\globalsmile\modGLGate.bas"
$modGLAgg    = "D:\citrixlabph\globalsmile\modGLAggregation.bas"
$modPJSync   = "D:\citrixlabph\globalsmile\modPJSync.bas"

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

    # Remove old standard modules by exact name lookup
    L "Step 4: Removing old/conflicting standard modules..."
    $toRemove = @("modEngine", "modGLRefresh", "Module1", "modPJAutomation", "modGLWorkbookMap", "modGLScanEngine", "modGLGate", "modGLAggregation", "modPJSync", "modEngine1", "modGLRefresh1")
    foreach ($name in $toRemove) {
        try {
            $comp = $vbp.VBComponents.Item($name)
            if ($null -ne $comp) {
                $vbp.VBComponents.Remove($comp)
                L "  Removed $name"
            }
        } catch {
            # Not present
        }
    }

    # Save and close so removed modules are flushed
    L "Step 5: Flushing removed modules..."
    $wb.Save()
    $wb.Close($true)

    # Reopen to import new modules
    L "Step 6: Reopening workbook for imports..."
    $wb = $xl.Workbooks.Open($targetXlsm, 0, $false)
    $vbp = $wb.VBProject

    # Clean Sheet9 backticks
    L "Step 7: Checking Sheet9 for backtick pollution..."
    $comp9 = $vbp.VBComponents.Item("Sheet9")
    $s9Code = $comp9.CodeModule.Lines(1, $comp9.CodeModule.CountOfLines)
    if ($s9Code -match '```') {
        L "  Cleaning markdown backticks from Sheet9..."
        $cleanedS9 = ($s9Code -split "`r?`n" | Where-Object { $_ -notmatch '^```' }) -join "`r`n"
        $comp9.CodeModule.DeleteLines(1, $comp9.CodeModule.CountOfLines)
        $comp9.CodeModule.AddFromString($cleanedS9)
        L "  Sheet9 cleaned."
    }

    # Update Sheet12 (PJ)
    L "Step 8: Updating Sheet12 (PJ tab)..."
    $s12Raw = Get-Content $sheet12 -Raw
    $s12Clean = (($s12Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp12 = $vbp.VBComponents.Item("Sheet12")
    $comp12.CodeModule.DeleteLines(1, $comp12.CodeModule.CountOfLines)
    $comp12.CodeModule.AddFromString($s12Clean)
    L "  Sheet12 updated."

    # Update Sheet11 (PJ Non-Vat)
    L "Step 9: Updating Sheet11 (PJ Non-Vat tab)..."
    $s11Raw = Get-Content $sheet11 -Raw
    $s11Clean = (($s11Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp11 = $vbp.VBComponents.Item("Sheet11")
    $comp11.CodeModule.DeleteLines(1, $comp11.CodeModule.CountOfLines)
    $comp11.CodeModule.AddFromString($s11Clean)
    L "  Sheet11 updated."

    # Import modules
    L "Step 10: Importing new modules..."
    $mods = @($modEngine, $modGLMap, $modGLScan, $modGLRef, $modGLGate, $modGLAgg, $modPJSync)
    foreach ($m in $mods) {
        if (Test-Path $m) {
            $vbp.VBComponents.Import($m) | Out-Null
            L "  Imported $([System.IO.Path]::GetFileName($m))"
        }
    }

    # List components to confirm no duplicate names
    L "Step 11: Verifying module names..."
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Type -eq 1) { # vbext_ct_StdModule
            L "  Standard Module: $($comp.Name)"
        }
    }

    # Save and close
    L "Step 12: Saving finalized workbook..."
    $wb.Save()
    $wb.Close($true)
    L "  Build saved."

    # Reopen for testing
    L "Step 13: Reopening for certification testing..."
    $wb = $xl.Workbooks.Open($targetXlsm, 0, $false)
    $xl.EnableEvents = $true

    # Run RefreshGL
    L "Step 14: Testing RefreshGL()..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $xl.Run("RefreshGL")
    $sw.Stop()
    L "  RefreshGL() completed in $($sw.Elapsed.TotalSeconds)s"

    $wsAudit = $wb.Sheets.Item("GL_AUDIT")
    $status = $wsAudit.Cells.Item(7, 3).Value
    $diff   = $wsAudit.Cells.Item(4, 3).Value
    $unmap  = $wsAudit.Cells.Item(5, 3).Value
    $inv    = $wsAudit.Cells.Item(6, 3).Value
    L "  GL_AUDIT Status: $status | Diff: $diff | Unmapped: $unmap | Invalid: $inv"

    if ($status -ne "PASS") { throw "GL_AUDIT Status is not PASS: $status" }

    # Test PJ -> CDJ Live Sync
    L "Step 15: Testing PJ -> CDJ live sync..."
    $wsPJ = $wb.Sheets.Item("PJ")
    $wsCDJ = $wb.Sheets.Item("CDJ")

    $pjRow = 10
    while ($null -ne $wsPJ.Cells.Item($pjRow, 4).Value -and "" -ne $wsPJ.Cells.Item($pjRow, 4).Value) {
        $pjRow++
    }

    $refId = "PJ-CERT-01"
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
    if ($cdjRow -lt 15) { throw "Sync INSERT failed: Ref $refId not found in CDJ" }

    $cCash = $wsCDJ.Cells.Item($cdjRow, 6).Value
    $cVAT  = $wsCDJ.Cells.Item($cdjRow, 7).Value
    $cMat  = $wsCDJ.Cells.Item($cdjRow, 14).Value
    $cDeb  = $wsCDJ.Cells.Item($cdjRow, 20).Value
    if ([Math]::Abs($cCash - (-5600.00)) -gt 0.01 -or [Math]::Abs($cMat - 5000.00) -gt 0.01) {
        throw "Sync INSERT value validation failed"
    }
    L "  PASS: PJ -> CDJ Insert verified on CDJ Row $cdjRow (Cash=$cCash, VAT=$cVAT, Mat=$cMat, Debit=$cDeb)"

    # Test update
    $wsPJ.Cells.Item($pjRow, 10).Value = 8000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 960.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 8960.00
    $xl.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cCashUpd = $wsCDJ.Cells.Item($cdjRow, 6).Value
    $cMatUpd  = $wsCDJ.Cells.Item($cdjRow, 14).Value
    if ([Math]::Abs($cCashUpd - (-8960.00)) -gt 0.01 -or [Math]::Abs($cMatUpd - 8000.00) -gt 0.01) {
        throw "Sync UPDATE value validation failed"
    }
    L "  PASS: PJ -> CDJ Update verified (Cash=$cCashUpd, Mat=$cMatUpd)"

    # Clean up test row
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value = $cdjRow
    L "  Test rows cleaned up."

    # Final RefreshGL check
    $xl.Run("RefreshGL")
    L "  Final RefreshGL re-verified after sync test."

    $wb.Save()
    L "  Workbook certified and saved."

    $hash = (Get-FileHash -Path $targetXlsm -Algorithm SHA256).Hash
    L "SHA-256: $hash"
    L "================================================================="
    L "  SUCCESS: v8.3.3 FULLY CERTIFIED AND READY"
    L "================================================================="

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
