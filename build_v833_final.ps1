# build_v833_final.ps1
# Complete clean build, PJ->CDJ sync verification, and GL certification for v8.3.3

$ErrorActionPreference = 'Stop'

$src = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3-FINAL.xlsm'
$dst = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm'
$log = 'D:\citrixlabph\globalsmile\build_v833_final.log'

"" | Out-File $log
function L($m) { $ts = Get-Date -f "HH:mm:ss"; "$ts  $m" | Tee-Object -Append -FilePath $log; Write-Host "$ts  $m" }

L "================================================================="
L "  BUILD & CERTIFY: Global-Smile_2026-v8.3.3.xlsm"
L "================================================================="

# 1. Copy base workbook
L "Step 1: Copying base workbook..."
Copy-Item $src $dst -Force
L "  Copied $src -> $dst"

# 2. Add-Type for background MsgBox auto-dismisser
L "Step 2: Starting background AutoDismisser..."
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

public class AutoDismisser {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);

  public static void Start(string logPath, int seconds) {
    Thread t = new Thread(delegate() {
      DateTime end = DateTime.UtcNow.AddSeconds(seconds);
      while (DateTime.UtcNow < end) {
        EnumWindows(delegate(IntPtr h, IntPtr l) {
          StringBuilder cn = new StringBuilder(64);
          GetClassName(h, cn, 64);
          if (cn.ToString() == "#32770") {
            StringBuilder tb = new StringBuilder(1024);
            GetWindowText(h, tb, 1024);
            string title = tb.ToString();
            if (title.Contains("SCAN") || title.Contains("RefreshGL") || title.Contains("Microsoft Excel") || title.Contains("PNV Automation") || title.Contains("PJ Sync") || title.Contains("Status")) {
              System.Collections.Generic.List<string> lines = new System.Collections.Generic.List<string>();
              lines.Add("MSGBOX: [" + title + "]");
              EnumChildWindows(h, delegate(IntPtr k, IntPtr l2) {
                StringBuilder kb = new StringBuilder(2048);
                GetWindowText(k, kb, 2048);
                if (kb.Length > 0) lines.Add("   text: " + kb.ToString());
                return true;
              }, IntPtr.Zero);
              File.AppendAllLines(logPath, lines);
              SendMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
            }
          }
          return true;
        }, IntPtr.Zero);
        Thread.Sleep(200);
      }
    });
    t.IsBackground = true;
    t.Start();
  }
}
'@

[AutoDismisser]::Start($log, 60)
L "  AutoDismisser active."

# 3. Open Excel COM
L "Step 3: Launching Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.AutomationSecurity = 1 # msoAutomationSecurityLow

$wb = $null
try {
    $wb = $excel.Workbooks.Open($dst, 0, $false)
    $vbp = $wb.VBProject
    L "  Workbook opened in Excel COM."

    # 4. Clean Sheet9 (remove markdown backticks)
    L "Step 4: Cleaning Sheet9..."
    try {
        $s9 = $vbp.VBComponents.Item("Sheet9")
        if ($s9 -ne $null) {
            $cm9 = $s9.CodeModule
            $code9 = $cm9.Lines(1, $cm9.CountOfLines)
            $targetBacktick = '```'
            if ($code9.Contains($targetBacktick)) {
                $code9 = $code9.Replace($targetBacktick, '')
                $cm9.DeleteLines(1, $cm9.CountOfLines)
                $cm9.AddFromString($code9)
                L "  Cleaned markdown backticks from Sheet9."
            }
        }
    } catch {
        L "  Warning on Sheet9 cleaning: $($_.Exception.Message)"
    }

    # 5. Update Sheet12 (PJ tab code)
    L "Step 5: Updating Sheet12 (PJ tab change handler)..."
    $s12Src = "D:\citrixlabph\globalsmile\Sheet12_v833_hardened.cls"
    if (Test-Path $s12Src) {
        $s12Raw = Get-Content $s12Src -Raw
        $s12Clean = (($s12Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
        $comp12 = $vbp.VBComponents.Item("Sheet12")
        $comp12.CodeModule.DeleteLines(1, $comp12.CodeModule.CountOfLines)
        $comp12.CodeModule.AddFromString($s12Clean)
        L "  Sheet12 updated."
    }

    # 6. Update Sheet11 (PJ Non-Vat tab code)
    L "Step 6: Updating Sheet11 (PJ Non-Vat tab change handler)..."
    $s11Src = "D:\citrixlabph\globalsmile\Sheet11_v833_hardened.cls"
    if (Test-Path $s11Src) {
        $s11Raw = Get-Content $s11Src -Raw
        $s11Clean = (($s11Raw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
        $comp11 = $vbp.VBComponents.Item("Sheet11")
        $comp11.CodeModule.DeleteLines(1, $comp11.CodeModule.CountOfLines)
        $comp11.CodeModule.AddFromString($s11Clean)
        L "  Sheet11 updated."
    }

    # 7. Update modEngine in place with hardened version
    L "Step 7: Updating modEngine in place with hardened code..."
    $meSrc = "D:\citrixlabph\globalsmile\modEngine_v833_hardened.bas"
    $meRaw = Get-Content $meSrc -Raw
    $meClean = (($meRaw -split "`r?`n" | Where-Object { $_ -notmatch "^Attribute\s+" }) -join "`r`n")
    $compME = $vbp.VBComponents.Item("modEngine")
    $compME.CodeModule.DeleteLines(1, $compME.CodeModule.CountOfLines)
    $compME.CodeModule.AddFromString($meClean)
    L "  modEngine updated with hardened PJ->CDJ sync fixes."

    # 8. Remove legacy conflicting modules
    L "Step 8: Removing legacy modules..."
    $modsToRemove = @("modGLRefresh", "modGLScanEngine", "modGLWorkbookMap", "modGLGate", "modPJSync", "Module1", "modPJAutomation")
    foreach ($m in $modsToRemove) {
        try {
            $comp = $vbp.VBComponents.Item($m)
            if ($comp -ne $null) {
                $vbp.VBComponents.Remove($comp)
                L "  Removed $m"
            }
        } catch {}
    }

    # 9. Import updated .bas modules
    L "Step 9: Importing verified standard modules..."
    $modsToImport = @(
        'D:\citrixlabph\globalsmile\modGLWorkbookMap.bas',
        'D:\citrixlabph\globalsmile\modGLScanEngine.bas',
        'D:\citrixlabph\globalsmile\modGLRefresh.bas',
        'D:\citrixlabph\globalsmile\modGLGate.bas',
        'D:\citrixlabph\globalsmile\modPJSync.bas'
    )
    foreach ($p in $modsToImport) {
        $vbp.VBComponents.Import($p) | Out-Null
        L "  Imported $([System.IO.Path]::GetFileName($p))"
    }

    # 10. Run RefreshGL macro to verify runtime execution
    L "Step 10: Running RefreshGL()..."
    [AutoDismisser]::Start($log, 30)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run("RefreshGL")
    $sw.Stop()
    L "  RefreshGL completed in $($sw.Elapsed.TotalSeconds) seconds."

    # 11. Verify GL Results
    L "Step 11: Verifying GL Balances..."
    $gl = $wb.Worksheets.Item("GL")
    $calc = $wb.Worksheets.Item("GL_V8_CALC")
    $audit = $wb.Worksheets.Item("GL_AUDIT")

    $miscJanDebit = $gl.Cells.Item(520, 7).Value2
    $miscJanCredit = $gl.Cells.Item(520, 8).Value2
    $miscJanBal = $gl.Cells.Item(520, 9).Value2
    L "  GL!Miscellaneous Jan (row 520): Debit=$miscJanDebit, Credit=$miscJanCredit, EndingBal=$miscJanBal"

    $cibJanDebit = $gl.Cells.Item(13, 7).Value2
    $cibJanCredit = $gl.Cells.Item(13, 8).Value2
    $cibJanBal = $gl.Cells.Item(13, 9).Value2
    L "  GL!Cash in Bank Jan (row 13): Debit=$cibJanDebit, Credit=$cibJanCredit, EndingBal=$cibJanBal"

    $calcRowCount = $calc.UsedRange.Rows.Count
    L "  GL_V8_CALC rows: $calcRowCount"

    $auditStatus = $audit.Cells.Item(7, 2).Value2
    $auditSummary = $audit.Cells.Item(7, 3).Value2
    L "  GL_AUDIT Status: $auditStatus | Summary: $auditSummary"

    if ($auditStatus -ne "PASS") { throw "Audit gate failed: $auditStatus - $auditSummary" }

    # 12. Live PJ -> CDJ Sync Verification Test
    L "Step 12: Running Live PJ -> CDJ Sync Verification..."
    $wsPJ = $wb.Worksheets.Item("PJ")
    $wsCDJ = $wb.Worksheets.Item("CDJ")

    $pjRow = 10
    while ($null -ne $wsPJ.Cells.Item($pjRow, 4).Value2 -and "" -ne $wsPJ.Cells.Item($pjRow, 4).Value2) {
        $pjRow++
    }
    L "  Using test row $pjRow in PJ..."

    $testRef = "PJ-TEST-833"
    $wsPJ.Cells.Item($pjRow, 2).Value = "JAN"
    $wsPJ.Cells.Item($pjRow, 4).Value = "233-251-708"
    $wsPJ.Cells.Item($pjRow, 5).Value = "FEDERAL BRENT RETAIL INC"
    $wsPJ.Cells.Item($pjRow, 7).Value = "Clinic Materials and Supplies"
    $wsPJ.Cells.Item($pjRow, 8).Value = $testRef
    $wsPJ.Cells.Item($pjRow, 10).Value = 5000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 600.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 5600.00

    # Call sync
    $excel.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)
    $cdjRow = $excel.Run("FindCDJRow", $testRef)
    L "  CDJ Synced Row: $cdjRow"

    if ($cdjRow -lt 15) { throw "PJ->CDJ sync test failed: row not created" }

    $cCash = $wsCDJ.Cells.Item($cdjRow, 6).Value2
    $cVAT  = $wsCDJ.Cells.Item($cdjRow, 7).Value2
    $cMat  = $wsCDJ.Cells.Item($cdjRow, 14).Value2
    $cDeb  = $wsCDJ.Cells.Item($cdjRow, 20).Value2
    L "  CDJ Synced Values: Cash=$cCash, VAT=$cVAT, ClinicMat=$cMat, Debit=$cDeb"

    if ([Math]::Abs($cCash - (-5600.00)) -gt 0.01 -or [Math]::Abs($cMat - 5000.00) -gt 0.01) {
        throw "PJ->CDJ sync test value mismatch"
    }
    L "  [PASS] PJ -> CDJ INSERT verified!"

    # Test Update on same row
    $wsPJ.Cells.Item($pjRow, 10).Value = 8000.00
    $wsPJ.Cells.Item($pjRow, 13).Value = 960.00
    $wsPJ.Cells.Item($pjRow, 14).Value = 8960.00
    $excel.Run("SyncPJToCDJ", $wsPJ, [long]$pjRow)

    $cCashUpd = $wsCDJ.Cells.Item($cdjRow, 6).Value2
    $cMatUpd  = $wsCDJ.Cells.Item($cdjRow, 14).Value2
    L "  CDJ Values After Update: Cash=$cCashUpd, ClinicMat=$cMatUpd"
    if ([Math]::Abs($cCashUpd - (-8960.00)) -gt 0.01 -or [Math]::Abs($cMatUpd - 8000.00) -gt 0.01) {
        throw "PJ->CDJ sync update value mismatch"
    }
    L "  [PASS] PJ -> CDJ UPDATE verified!"

    # Clean up test rows
    $wsPJ.Rows.Item($pjRow).ClearContents() | Out-Null
    $wsCDJ.Rows.Item($cdjRow).ClearContents() | Out-Null
    $wsCDJ.Cells.Item(14, 27).Value2 = [string]$cdjRow
    L "  Test row cleaned up."

    # Final RefreshGL to leave GL 100% clean
    L "Step 13: Running final RefreshGL()..."
    [AutoDismisser]::Start($log, 30)
    $excel.Run("RefreshGL")
    L "  Final RefreshGL finished."

    # Save and close workbook before hashing
    L "Step 14: Saving certified workbook..."
    $wb.Save()
    $wb.Close($true)
    $wb = $null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    $excel = $null
    L "  Workbook saved and Excel closed."

    # Hash
    $hash = (Get-FileHash -Path $dst -Algorithm SHA256).Hash
    L "SHA-256: $hash"
    L "================================================================="
    L "  SUCCESS: v8.3.3 FULLY CERTIFIED & PRODUCTION READY"
    L "================================================================="

} finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
