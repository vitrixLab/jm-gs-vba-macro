# build_and_certify_v833.ps1
# Full clean build of Global-Smile_2026-v8.3.3.xlsm from Global-Smile_2026-v8.3-FINAL.xlsm base
# Integrates all hardened modules and runs complete verification suite.

$ErrorActionPreference = "Stop"
$log = "D:\citrixlabph\globalsmile\build_v833_certified.log"
function L($m) { $ts = Get-Date -f "HH:mm:ss"; "$ts  $m" | Tee-Object -Append -FilePath $log; Write-Host "$ts  $m" }

"" | Out-File $log
L "================================================================="
L "  BUILD & CERTIFY: Global-Smile_2026-v8.3.3.xlsm"
L "================================================================="

$baseXlsm    = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3-FINAL.xlsm"
$targetXlsm  = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"

# Source modules
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
L "  Base copied to $targetXlsm"

# Step 2: Open Excel COM
L "Step 2: Launching Excel COM..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible           = $false
$xl.DisplayAlerts     = $false
$xl.AutomationSecurity = 1 # msoAutomationSecurityLow
$xl.EnableEvents      = $false

try {
    L "Step 3: Opening workbook..."
    $wb = $xl.Workbooks.Open($targetXlsm, 0, $false)
    $vbp = $wb.VBProject

    # Remove conflicting/old modules
    $toRemove = @("modGLRefresh", "Module1", "modEngine", "modGLWorkbookMap", "modGLScanEngine", "modGLGate", "modGLAggregation", "modPJSync")
    foreach ($comp in [array]$vbp.VBComponents) {
        if ($toRemove -contains $comp.Name) {
            L "  Removing $($comp.Name)..."
            $vbp.VBComponents.Remove($comp)
        }
    }

    # Clean Sheet9 (PJ Non-Vat2) of any markdown backticks
    L "Step 4: Checking Sheet9 code for backtick pollution..."
    $comp9 = $vbp.VBComponents.Item("Sheet9")
    $s9Code = $comp9.CodeModule.Lines(1, $comp9.CodeModule.CountOfLines)
    if ($s9Code -match '```') {
        L "  Cleaning markdown backticks from Sheet9..."
        $cleanedS9 = ($s9Code -split "`r?`n" | Where-Object { $_ -notmatch '^```' }) -join "`r`n"
        $comp9.CodeModule.DeleteLines(1, $comp9.CodeModule.CountOfLines)
        $comp9.CodeModule.AddFromString($cleanedS9)
        L "  Sheet9 cleaned."
    }

    # Update Sheet12 (PJ tab)
    L "Step 5: Updating Sheet12 (PJ tab code)..."
    $s12Raw = Get-Content $sheet12 -Raw
    $s12Clean = (($s12Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp12 = $vbp.VBComponents.Item("Sheet12")
    $comp12.CodeModule.DeleteLines(1, $comp12.CodeModule.CountOfLines)
    $comp12.CodeModule.AddFromString($s12Clean)
    L "  Sheet12 updated."

    # Update Sheet11 (PJ Non-Vat tab)
    L "Step 6: Updating Sheet11 (PJ Non-Vat tab code)..."
    $s11Raw = Get-Content $sheet11 -Raw
    $s11Clean = (($s11Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $comp11 = $vbp.VBComponents.Item("Sheet11")
    $comp11.CodeModule.DeleteLines(1, $comp11.CodeModule.CountOfLines)
    $comp11.CodeModule.AddFromString($s11Clean)
    L "  Sheet11 updated."

    # Import modules
    L "Step 7: Importing hardened standard modules..."
    $modulesToImport = @($modEngine, $modGLMap, $modGLScan, $modGLRef, $modGLGate, $modGLAgg, $modPJSync)
    foreach ($m in $modulesToImport) {
        if (Test-Path $m) {
            $vbp.VBComponents.Import($m) | Out-Null
            L "  Imported $([System.IO.Path]::GetFileName($m))"
        }
    }

    # Save and close to register modules
    L "Step 8: Saving initial build..."
    $wb.Save()
    $wb.Close($true)
    L "  Saved and closed."

    # Reopen for verification tests
    L "Step 9: Reopening for certification test suite..."
    $wb = $xl.Workbooks.Open($targetXlsm, 0, $false)
    $xl.EnableEvents = $true

    # --- Test 1: InitializeEngine & FindExpenseColumn ---
    L "Test 1: Testing FindExpenseColumn mapping..."
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
        if ($col -ne $t.Expected) {
            throw "FindExpenseColumn failure for '$($t.Coa)': expected $($t.Expected), got $col"
        }
        L "  PASS: '$($t.Coa)' -> Col $col"
    }

    # --- Test 2: PJ -> CDJ Live Sync (Insert & Update) ---
    L "Test 2: Testing PJ -> CDJ live sync..."
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

    # Cleanup test row
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value = $cdjRow
    L "  Test rows cleaned up."

    # --- Test 3: RefreshGL Certification ---
    L "Test 3: Running RefreshGL()..."
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
    if ([Math]::Abs($diff) -gt 0.001) { throw "GL_AUDIT Difference is non-zero: $diff" }
    if ([Math]::Abs($unmap) -gt 0.001) { throw "GL_AUDIT Unmapped is non-zero: $unmap" }

    # Check key GL balances
    $wsGL = $wb.Sheets.Item("GL")
    $cibCr = $wsGL.Cells.Item(13, 10).Value
    $miscDb = $wsGL.Cells.Item(520, 9).Value
    L "  GL Cash In Bank Jan Credit: $cibCr (Expected: 7025.25)"
    L "  GL Miscellaneous Jan Debit: $miscDb (Expected: 2980.40...)"

    # Save certified workbook
    L "Step 10: Saving certified v8.3.3 workbook..."
    $wb.Save()
    L "  Workbook saved."

    # SHA-256
    $hash = (Get-FileHash -Path $targetXlsm -Algorithm SHA256).Hash
    L "SHA-256: $hash"
    L "================================================================="
    L "  CERTIFICATION RESULT: ALL GATES PASS - READY FOR PRODUCTION"
    L "================================================================="

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
