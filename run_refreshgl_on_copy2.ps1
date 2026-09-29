# run_refreshgl_on_copy2.ps1
# Runs the workbook's own RefreshGL on a SCRATCH COPY. The original file is only read.
# A compiled background thread auto-dismisses the RefreshGL MsgBox and logs its text.
$ErrorActionPreference = 'Stop'
$src = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.2-final-review.xlsm'
$dst = 'D:\citrixlabph\globalsmile\v8.2_review\scratch_run.xlsm'
$out = 'D:\citrixlabph\globalsmile\v8.2_review\runtime_refreshgl.txt'
$h0 = (Get-FileHash $src -Algorithm SHA256).Hash
Copy-Item $src $dst -Force
"original sha256 before: $h0" | Out-File $out -Encoding utf8

Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

public class Dismisser {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);

  public static void Start(string logPath, int seconds) {
    Thread t = new Thread(delegate() {
      DateTime end = DateTime.UtcNow.AddSeconds(seconds);
      int seen = 0;
      while (DateTime.UtcNow < end) {
        System.Collections.Generic.List<IntPtr> hits = new System.Collections.Generic.List<IntPtr>();
        System.Collections.Generic.List<string> titles = new System.Collections.Generic.List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr l) {
          StringBuilder cn = new StringBuilder(64);
          GetClassName(h, cn, 64);
          if (cn.ToString() == "#32770") {
            StringBuilder tb = new StringBuilder(1024);
            GetWindowText(h, tb, 1024);
            string title = tb.ToString();
            if (title.StartsWith("RefreshGL")) { hits.Add(h); titles.Add(title); }
          }
          return true;
        }, IntPtr.Zero);
        for (int i = 0; i < hits.Count; i++) {
          System.Collections.Generic.List<string> lines = new System.Collections.Generic.List<string>();
          lines.Add("MSGBOX CAPTION: [" + titles[i] + "]");
          EnumChildWindows(hits[i], delegate(IntPtr h, IntPtr l) {
            StringBuilder tb = new StringBuilder(2048);
            GetWindowText(h, tb, 2048);
            if (tb.Length > 0) lines.Add("   text: " + tb.ToString());
            return true;
          }, IntPtr.Zero);
          File.AppendAllLines(logPath, lines);
          SendMessage(hits[i], 0x0010, IntPtr.Zero, IntPtr.Zero);
          Thread.Sleep(300);
          SendMessage(hits[i], 0x0111, (IntPtr)1, IntPtr.Zero);
          seen++;
        }
        Thread.Sleep(250);
      }
      File.AppendAllText(logPath, "dismisser thread ended; dialogs dismissed=" + seen + Environment.NewLine);
    });
    t.IsBackground = true;
    t.Start();
  }
}
'@

[Dismisser]::Start($out, 240)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.AutomationSecurity = 1
$wb = $null
$excelPid = 0
try {
    $wb = $excel.Workbooks.Open($dst, 0, $false)
    $excelPid = (Get-Process EXCEL | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    "opened scratch copy: $($wb.Name) (excel pid $excelPid)" | Out-File $out -Append -Encoding utf8
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run("RefreshGL")
    $sw.Stop()
    "RefreshGL returned after $([math]::Round($sw.Elapsed.TotalSeconds,1)) s" | Out-File $out -Append -Encoding utf8

    $gl = $wb.Worksheets("GL")
    $lastTitleRow = 0
    for ($r = 13; $r -le 1200; $r++) {
        if ($gl.Cells($r, 6).Value2 -ne $null) { $lastTitleRow = $r }
    }
    "GL last non-empty Account Title row = $lastTitleRow" | Out-File $out -Append -Encoding utf8
    "sample GL rows after refresh (B=month, E=label, F=account title, G=debit, H=credit, I=ending):" |
        Out-File $out -Append -Encoding utf8
    foreach ($r in 11,12,13,14,24,25,26,193,205,553,564,565,577,609,610) {
        "   row {0,4}  B={1,-6} C={2,-4} E={3,-16} F={4,-34} G={5,-12} H={6,-12} I={7}" -f $r,
          $gl.Cells($r,2).Value2, $gl.Cells($r,3).Value2, $gl.Cells($r,5).Value2, $gl.Cells($r,6).Value2,
          $gl.Cells($r,7).Value2, $gl.Cells($r,8).Value2, $gl.Cells($r,9).Value2 |
            Out-File $out -Append -Encoding utf8
    }
    $td = 0.0; $tc = 0.0; $nz = 0; $titles = 0
    for ($r = 13; $r -le $lastTitleRow; $r++) {
        $g = $gl.Cells($r,7).Value2; $h = $gl.Cells($r,8).Value2
        if ($g -is [double]) { $td += $g }
        if ($h -is [double]) { $tc += $h }
        if ((($g -is [double]) -and $g -ne 0) -or (($h -is [double]) -and $h -ne 0)) { $nz++ }
        if ($gl.Cells($r,6).Value2 -ne $null) { $titles++ }
    }
    "GL after refresh: title rows={0} debit={1:N2} credit={2:N2} difference={3:N2} active cells={4}" -f `
        $titles, $td, $tc, ($td - $tc), $nz | Out-File $out -Append -Encoding utf8
    $au = $wb.Worksheets("GL_AUDIT")
    "GL_AUDIT after refresh:" | Out-File $out -Append -Encoding utf8
    for ($r = 1; $r -le 4; $r++) {
        "   {0} | {1} | {2} | {3}" -f $au.Cells($r,1).Value2,$au.Cells($r,2).Value2,$au.Cells($r,3).Value2,$au.Cells($r,4).Value2 |
            Out-File $out -Append -Encoding utf8
    }
    $cv = $wb.Worksheets("GL_V8_CALC")
    "GL_V8_CALC row2 after refresh: {0} | {1} | {2} | {3} | {4}" -f `
        $cv.Cells(2,1).Value2,$cv.Cells(2,2).Value2,$cv.Cells(2,3).Value2,$cv.Cells(2,4).Value2,$cv.Cells(2,5).Value2 |
        Out-File $out -Append -Encoding utf8
    "GL_V8_CALC row553 after refresh: {0} | {1} | {2} | {3} | {4}" -f `
        $cv.Cells(553,1).Value2,$cv.Cells(553,2).Value2,$cv.Cells(553,3).Value2,$cv.Cells(553,4).Value2,$cv.Cells(553,5).Value2 |
        Out-File $out -Append -Encoding utf8
    "ValidateGLConsistency after refresh = $($excel.Run('ValidateGLConsistency'))" | Out-File $out -Append -Encoding utf8
}
finally {
    if ($wb -ne $null) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
Start-Sleep -Seconds 2
if ($excelPid -gt 0) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
$h1 = (Get-FileHash $src -Algorithm SHA256).Hash
"original sha256 after: $h1" | Out-File $out -Append -Encoding utf8
"original unchanged: $($h0 -eq $h1)" | Out-File $out -Append -Encoding utf8
Get-ChildItem 'D:\citrixlabph\globalsmile' -Filter '~$*' -Force -ErrorAction SilentlyContinue |
    ForEach-Object { "lock file removed: $($_.Name)"; Remove-Item $_.FullName -Force }
"done" | Out-File $out -Append -Encoding utf8
Get-Content $out

