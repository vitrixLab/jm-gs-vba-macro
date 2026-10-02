# build_v833.ps1 - Builds and tests Global-Smile_2026-v8.3.3.xlsm
$ErrorActionPreference = 'Stop'

$src = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3-FINAL.xlsm'
$dst = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm'
$log = 'D:\citrixlabph\globalsmile\build_v833.log'

Write-Output "Starting v8.3.3 build process..." | Tee-Object -FilePath $log

# 1. Copy base workbook
Copy-Item $src $dst -Force
Write-Output "Copied $src -> $dst" | Tee-Object -FilePath $log -Append

# 2. Add-Type for background MsgBox auto-dismisser
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
            if (title.Contains("SCAN") || title.Contains("RefreshGL") || title.Contains("Microsoft Excel") || title.Contains("PNV Automation")) {
              System.Collections.Generic.List<string> lines = new System.Collections.Generic.List<string>();
              lines.Add("MSGBOX CAPTION: [" + title + "]");
              EnumChildWindows(h, delegate(IntPtr k, IntPtr l2) {
                StringBuilder kb = new StringBuilder(2048);
                GetWindowText(k, kb, 2048);
                if (kb.Length > 0) lines.Add("   text: " + kb.ToString());
                return true;
              }, IntPtr.Zero);
              File.AppendAllLines(logPath, lines);
              SendMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
              Thread.Sleep(200);
              SendMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero);  // IDOK
            }
          }
          return true;
        }, IntPtr.Zero);
        Thread.Sleep(250);
      }
    });
    t.IsBackground = true;
    t.Start();
  }
}
'@

[AutoDismisser]::Start($log, 180)

# 3. Open workbook via Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.AutomationSecurity = 1 # Allow macros

$wb = $null
try {
    $wb = $excel.Workbooks.Open($dst, 0, $false)
    Write-Output "Workbook opened in Excel COM." | Tee-Object -FilePath $log -Append

    # 4. Clean Sheet9 (remove markdown backticks using single-quoted strings)
    try {
        $s9 = $wb.VBProject.VBComponents.Item("Sheet9")
        if ($s9 -ne $null) {
            $cm9 = $s9.CodeModule
            $code9 = $cm9.Lines(1, $cm9.CountOfLines)
            $targetBacktick = '```'
            if ($code9.Contains($targetBacktick)) {
                $code9 = $code9.Replace($targetBacktick, '')
                $cm9.DeleteLines(1, $cm9.CountOfLines)
                $cm9.AddFromString($code9)
                Write-Output "Cleaned markdown backticks from Sheet9." | Tee-Object -FilePath $log -Append
            }
        }
    } catch {
        Write-Output "Warning on Sheet9 cleaning: $($_.Exception.Message)" | Tee-Object -FilePath $log -Append
    }

    # 5. Neutralize colliding public procedures in modEngine
    try {
        $me = $wb.VBProject.VBComponents.Item("modEngine")
        if ($me -ne $null) {
            $cme = $me.CodeModule
            $lines = $cme.Lines(1, $cme.CountOfLines)
            $lines = $lines -replace '(?m)^(\s*)Public Sub RefreshGL\(', '$1Private Sub Legacy_RefreshGL('
            $lines = $lines -replace '(?m)^(\s*)Public Sub RefreshAllGL\(', '$1Private Sub Legacy_RefreshAllGL('
            $lines = $lines -replace '(?m)^(\s*)Public Function CalculateMonthlyNet\(', '$1Private Function Legacy_CalculateMonthlyNet('
            $lines = $lines -replace '(?m)^(\s*)Public Function CalculateEndingBalanceFull\(', '$1Private Function Legacy_CalculateEndingBalanceFull('
            $lines = $lines -replace '(?m)^(\s*)Public Function ValidateGLConsistency\(', '$1Private Function Legacy_ValidateGLConsistency('
            $cme.DeleteLines(1, $cme.CountOfLines)
            $cme.AddFromString($lines)
            Write-Output "Neutralized colliding legacy procedures in modEngine." | Tee-Object -FilePath $log -Append
        }
    } catch {
        Write-Output "Warning on modEngine: $($_.Exception.Message)" | Tee-Object -FilePath $log -Append
    }

    # 6. Remove existing modules to be updated
    $modsToRemove = @("modGLRefresh", "modGLScanEngine", "modGLWorkbookMap", "modGLGate", "modPJSync", "Module1")
    foreach ($m in $modsToRemove) {
        try {
            $comp = $wb.VBProject.VBComponents.Item($m)
            if ($comp -ne $null) {
                $wb.VBProject.VBComponents.Remove($comp)
                Write-Output "Removed old component: $m" | Tee-Object -FilePath $log -Append
            }
        } catch {}
    }

    # 7. Import updated .bas modules
    $modsToImport = @(
        'D:\citrixlabph\globalsmile\modGLWorkbookMap.bas',
        'D:\citrixlabph\globalsmile\modGLScanEngine.bas',
        'D:\citrixlabph\globalsmile\modGLRefresh.bas',
        'D:\citrixlabph\globalsmile\modGLGate.bas',
        'D:\citrixlabph\globalsmile\modPJSync.bas'
    )
    foreach ($p in $modsToImport) {
        $wb.VBProject.VBComponents.Import($p)
        Write-Output "Imported component: $(Split-Path $p -Leaf)" | Tee-Object -FilePath $log -Append
    }

    # 8. Run RefreshGL macro to verify runtime execution
    Write-Output "Running RefreshGL()..." | Tee-Object -FilePath $log -Append
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run("RefreshGL")
    $sw.Stop()
    Write-Output "RefreshGL completed in $($sw.Elapsed.TotalSeconds) seconds." | Tee-Object -FilePath $log -Append

    # 9. Verify results
    $gl = $wb.Worksheets.Item("GL")
    $calc = $wb.Worksheets.Item("GL_V8_CALC")
    $audit = $wb.Worksheets.Item("GL_AUDIT")

    # Check Miscellaneous / Jan (Block 39, Row 520)
    $miscJanDebit = $gl.Cells.Item(520, 7).Value2
    $miscJanCredit = $gl.Cells.Item(520, 8).Value2
    $miscJanBal = $gl.Cells.Item(520, 9).Value2
    Write-Output "GL!Miscellaneous Jan (row 520): Debit=$miscJanDebit, Credit=$miscJanCredit, EndingBal=$miscJanBal" | Tee-Object -FilePath $log -Append

    # Check Cash in Bank / Jan (Block 0, Row 13)
    $cashJanDebit = $gl.Cells.Item(13, 7).Value2
    $cashJanCredit = $gl.Cells.Item(13, 8).Value2
    $cashJanBal = $gl.Cells.Item(13, 9).Value2
    Write-Output "GL!Cash in Bank Jan (row 13): Debit=$cashJanDebit, Credit=$cashJanCredit, EndingBal=$cashJanBal" | Tee-Object -FilePath $log -Append

    # Check GL_V8_CALC row count
    $lastCalcRow = $calc.Cells.Item($calc.Rows.Count, 1).End(-4162).Row # xlUp
    $calcCount = $lastCalcRow - 1
    Write-Output "GL_V8_CALC data rows: $calcCount (expected 552)" | Tee-Object -FilePath $log -Append

    # Check GL_AUDIT status
    $lastAuditRow = $audit.Cells.Item($audit.Rows.Count, 1).End(-4162).Row
    $auditStatus = $audit.Cells.Item($lastAuditRow, 2).Value2
    $auditSummary = $audit.Cells.Item($lastAuditRow, 3).Value2
    Write-Output "GL_AUDIT row $($lastAuditRow): Status=$auditStatus, Summary=$auditSummary" | Tee-Object -FilePath $log -Append

    # 10. Save workbook
    $wb.Save()
    Write-Output "Workbook saved successfully to $dst" | Tee-Object -FilePath $log -Append

} catch {
    Write-Output "ERROR during build: $($_.Exception.Message)" | Tee-Object -FilePath $log -Append
    throw $_
} finally {
    if ($wb -ne $null) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

$hash = (Get-FileHash $dst -Algorithm SHA256).Hash
Write-Output "Build complete! SHA256 of $($dst): $hash" | Tee-Object -FilePath $log -Append
