# run_refreshgl_on_copy3.ps1
# Second-order demo: unmerge the GL header band (the only thing blocking RefreshGL),
# then run RefreshGL on a SCRATCH COPY and measure what it actually produces.
$ErrorActionPreference = 'Stop'
$src = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.2-final-review.xlsm'
$dst = 'D:\citrixlabph\globalsmile\v8.2_review\scratch_run2.xlsm'
$out = 'D:\citrixlabph\globalsmile\v8.2_review\runtime_refreshgl_unmerged.txt'
$h0 = (Get-FileHash $src -Algorithm SHA256).Hash
Copy-Item $src $dst -Force
"original sha256 before: $h0" | Out-File $out -Encoding utf8

Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
public class Dis2 {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
  public static void Start(string logPath, int seconds) {
    Thread t = new Thread(delegate() {
      DateTime end = DateTime.UtcNow.AddSeconds(seconds); int seen = 0;
      while (DateTime.UtcNow < end) {
        System.Collections.Generic.List<IntPtr> hits = new System.Collections.Generic.List<IntPtr>();
        System.Collections.Generic.List<string> titles = new System.Collections.Generic.List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr l) {
          StringBuilder cn = new StringBuilder(64); GetClassName(h, cn, 64);
          if (cn.ToString() == "#32770") {
            StringBuilder tb = new StringBuilder(1024); GetWindowText(h, tb, 1024);
            string title = tb.ToString();
            bool want = title.StartsWith("RefreshGL") || title == "Microsoft Excel";
            if (want) {
              System.Collections.Generic.List<string> kids = new System.Collections.Generic.List<string>();
              EnumChildWindows(h, delegate(IntPtr k, IntPtr l2) {
                StringBuilder kb = new StringBuilder(4096); GetWindowText(k, kb, 4096);
                if (kb.Length > 0) kids.Add(kb.ToString());
                return true;
              }, IntPtr.Zero);
              bool msg = false;
              foreach (string k in kids) { if (k.IndexOf("GL refreshed") >= 0 || k.IndexOf("RefreshGL") >= 0) msg = true; }
              if (msg) { hits.Add(h); titles.Add(title); }
            }
          }
          return true;
        }, IntPtr.Zero);
        for (int i = 0; i < hits.Count; i++) {
          System.Collections.Generic.List<string> lines = new System.Collections.Generic.List<string>();
          lines.Add("MSGBOX CAPTION: [" + titles[i] + "]");
          EnumChildWindows(hits[i], delegate(IntPtr h, IntPtr l) {
            StringBuilder tb = new StringBuilder(2048); GetWindowText(h, tb, 2048);
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
    t.IsBackground = true; t.Start();
  }
}
'@

[Dis2]::Start($out, 240)
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
    $gl = $wb.Worksheets("GL")
    # unmerge only the header band that RefreshGL writes as an array
    $unmerged = @()
    foreach ($m in @($gl.Range("B11:I12").MergeArea)) { }
    for ($r = 11; $r -le 12; $r++) {
        for ($c = 2; $c -le 9; $c++) {
            $cell = $gl.Cells($r, $c)
            if ($cell.MergeCells) {
                $area = $cell.MergeArea.Address(0, 0)
                if (-not ($unmerged -contains $area)) { $unmerged += $area }
            }
        }
    }
    foreach ($a in $unmerged) { $gl.Range($a).UnMerge() }
    "unmerged on scratch copy: $($unmerged -join ', ')" | Out-File $out -Append -Encoding utf8

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run("RefreshGL")
    $sw.Stop()
    "RefreshGL returned after $([math]::Round($sw.Elapsed.TotalSeconds,1)) s" | Out-File $out -Append -Encoding utf8

    "--- GL after refresh ---" | Out-File $out -Append -Encoding utf8
    foreach ($r in 11,12,13,14,24,25,26,205,553,564,565,566,577,609,610) {
        "   row {0,4}  B={1,-6} C={2,-4} E={3,-16} F={4,-34} G={5,-12} H={6,-12} I={7}" -f $r,
          $gl.Cells($r,2).Value2, $gl.Cells($r,3).Value2, $gl.Cells($r,5).Value2, $gl.Cells($r,6).Value2,
          $gl.Cells($r,7).Value2, $gl.Cells($r,8).Value2, $gl.Cells($r,9).Value2 |
            Out-File $out -Append -Encoding utf8
    }
    $titleRows = @()
    $td = 0.0; $tc = 0.0; $nz = 0
    for ($r = 13; $r -le 1200; $r++) {
        if ($gl.Cells($r,6).Value2 -ne $null) { $titleRows += $r }
        $g = $gl.Cells($r,7).Value2; $h = $gl.Cells($r,8).Value2
        if ($g -is [double]) { $td += $g }
        if ($h -is [double]) { $tc += $h }
        if ((($g -is [double]) -and $g -ne 0) -or (($h -is [double]) -and $h -ne 0)) { $nz++ }
    }
    "title rows written = {0}; first={1}; last={2}" -f $titleRows.Count, $titleRows[0], $titleRows[-1] |
        Out-File $out -Append -Encoding utf8
    "first 6 title rows: $($titleRows[0..5] -join ', ')  -> step={0}" -f `
        ($titleRows[1] - $titleRows[0]) | Out-File $out -Append -Encoding utf8
    "GL totals: debit={0:N2} credit={1:N2} difference={2:N2} active cells={3}" -f $td,$tc,($td-$tc),$nz |
        Out-File $out -Append -Encoding utf8
    $au = $wb.Worksheets("GL_AUDIT")
    "GL_AUDIT after refresh (rows 1-4):" | Out-File $out -Append -Encoding utf8
    for ($r = 1; $r -le 4; $r++) {
        "   {0} | {1} | {2} | {3}" -f $au.Cells($r,1).Value2,$au.Cells($r,2).Value2,$au.Cells($r,3).Value2,$au.Cells($r,4).Value2 |
            Out-File $out -Append -Encoding utf8
    }
    $cv = $wb.Worksheets("GL_V8_CALC")
    "GL_V8_CALC rows 2 and 553 after refresh: {0}|{1}|{2}|{3}|{4}   {5}|{6}|{7}|{8}|{9}" -f `
        $cv.Cells(2,1).Value2,$cv.Cells(2,2).Value2,$cv.Cells(2,3).Value2,$cv.Cells(2,4).Value2,$cv.Cells(2,5).Value2,
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
    ForEach-Object { Remove-Item $_.FullName -Force }
"done" | Out-File $out -Append -Encoding utf8
Get-Content $out

