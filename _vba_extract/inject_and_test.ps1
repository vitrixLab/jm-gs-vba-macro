$ErrorActionPreference = 'Stop'
$log = 'C:\citrixlabph\globalsmile\_vba_extract\inject_test.log'
$dlog = 'C:\citrixlabph\globalsmile\_vba_extract\dialog_watch.log'
$watch = 'C:\citrixlabph\globalsmile\_vba_extract\dialog_watch.ps1'
$basDir = 'C:\citrixlabph\globalsmile\_vba_extract\bas'
function L($m) { "$(Get-Date -Format 'HH:mm:ss.fff')  $m" | Out-File -FilePath $log -Append -Encoding utf8 }

$src = 'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.8.xlsm'
$dst = 'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm'
$bas = 'C:\citrixlabph\globalsmile\Graphify.bas'
$basCRLF = 'C:\citrixlabph\globalsmile\_vba_extract\Graphify_crlf.bas'

Remove-Item $log -ErrorAction SilentlyContinue
Remove-Item $dlog -ErrorAction SilentlyContinue
L "start"

# kill young EXCEL zombies (file locks); user sessions older than 30 min survive
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddMinutes(-30) } |
    ForEach-Object { L "killing stale EXCEL pid=$($_.Id)"; Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }

$watchProc = Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$watch,'-Log',$dlog -PassThru -WindowStyle Hidden
L "dialog watcher started (pid $($watchProc.Id))"

function ConvertTo-CrlfAscii([string]$p) {
    $t = [System.IO.File]::ReadAllText($p)
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace "`n", "`r`n"
    $out = [System.IO.Path]::ChangeExtension($p, '.crlf.bas')
    [System.IO.File]::WriteAllText($out, $t, [System.Text.Encoding]::ASCII)
    return $out
}

Copy-Item $src $dst -Force
L "copied v7.8 -> v7.9"

$xl = $null
$wb = $null
$ok = $false
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1
    L "excel $($xl.Version) started"

    $wb = $xl.Workbooks.Open($dst, 0, $false)
    L "workbook opened"
    $vbp = $wb.VBProject

    # ---------- inject fixed modules ----------
    $gBas = ConvertTo-CrlfAscii $bas
    foreach ($c in @($vbp.VBComponents)) { if ($c.Name -eq 'modGraphify') { $vbp.VBComponents.Remove($c) } }
    $comp = $vbp.VBComponents.Import($gBas)
    L "imported modGraphify lines=$($comp.CodeModule.CountOfLines)"

    foreach ($m in @('modEngine', 'Module1')) {
        $p = ConvertTo-CrlfAscii (Join-Path $basDir "$m.bas")
        foreach ($c in @($vbp.VBComponents)) { if ($c.Name -eq $m) { $vbp.VBComponents.Remove($c) } }
        $comp = $vbp.VBComponents.Import($p)
        L "imported $m lines=$($comp.CodeModule.CountOfLines)"
    }

    # ---------- Sheet9 (PJ Non-Vat module): replace code, keep designer binding
    $s9 = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq 'Sheet9') { $s9 = $c } }
    if ($s9 -eq $null) { throw 'Sheet9 component not found' }
    $s9txt = [System.IO.File]::ReadAllText((Join-Path $basDir 'Sheet9.cls'))
    $s9txt = $s9txt -replace "`r`n", "`n"
    $s9txt = ($s9txt -split "`n" | Where-Object { $_ -notmatch '^Attribute ' }) -join "`r`n"
    $s9txt = $s9txt.TrimStart("`r", "`n")
    $cm = $s9.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($s9txt)
    L "Sheet9 code replaced lines=$($cm.CountOfLines)"

    # ---------- verify injected markers ----------
    $meCode = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq 'modEngine') { $meCode = $c.CodeModule.Lines(1, $c.CodeModule.CountOfLines) } }
    L "modEngine CDJ_COL_DEBIT present: $($meCode -like '*Private Const CDJ_COL_DEBIT As Integer = 20*')"
    L "modEngine EXPENSE_END=19: $($meCode -like '*CDJ_EXPENSE_END As Integer = 19*')"
    L "modEngine T-write count: $([regex]::Matches($meCode, 'netAmount').Count) (expect >=4)"
    L "Sheet9 delegates: $($s9.CodeModule.Lines(1, $s9.CodeModule.CountOfLines) -like '*modEngine.SyncPJNonVatToCDJ Me, sourceRow*')"

    # ---------- compile-probe the edited writers on a throwaway sheet (no writes: ref=""/amount=0 -> exit)
    $tmp = $wb.Worksheets.Add()
    $tmp.Name = 'ZZZ_CompileProbe'
    $null = $xl.Run('SyncPJToCDJ', $tmp, 1)
    L "SyncPJToCDJ compiled + safe-exit OK"
    $null = $xl.Run('SyncPJNonVatToCDJ', $tmp, 1)
    L "SyncPJNonVatToCDJ compiled + safe-exit OK"
    $null = $xl.Run('CreateCDJFromPurchase', $tmp, 1)
    L "CreateCDJFromPurchase wrapper compiled OK"
    $tmp.Delete()
    L "probe sheet deleted"


    # ---------- CDJ: backfill T (Debit) for existing engine-created rows ----------
    # T = sum of category columns N..S + WHT (H) - additive only, T must be empty,
    # row must carry a reference in Y (PJ) or Z (PNV). U (Credit) stays manual.
    $cdj = $wb.Worksheets('CDJ')
    $nextRow = [int]$cdj.Range('AA14').Value2
    if ($nextRow -lt 15 -or $nextRow -gt 400) { $nextRow = 17 }
    $filled = 0
    for ($r = 15; $r -lt $nextRow; $r++) {
        $refY = $cdj.Range("Y$r").Text
        $refZ = $cdj.Range("Z$r").Text
        if (($refY -eq '' -and $refZ -eq '') -or $cdj.Range("T$r").Value2 -ne $null) { continue }
        $sum = 0.0
        foreach ($c in @('N','O','P','Q','R','S')) {
            $v = $cdj.Range("$c$r").Value2
            if ($v -ne $null) { $sum += [double]$v }
        }
        $h = $cdj.Range("H$r").Value2
        if ($h -ne $null) { $sum += [double]$h }
        if ($sum -ne 0) {
            $cdj.Range("T$r").Value2 = [math]::Round($sum, 2)
            $cdj.Range("T$r").NumberFormat = $cdj.Range("F$r").NumberFormat
            $filled++
            L "CDJ T$r = $sum (ref Y=$refY Z=$refZ)"
        }
    }
    L "CDJ Debit backfill filled=$filled rows"

    # ---------- Graphify validation ----------
    $pj = $wb.Worksheets('PJ')
    $pj.Activate()
    $pj.Range('A1').Select()
    $xl.Run('Graphify')
    L "PJ Graphify run #1 (create, auto-pick) OK"
    $co = $pj.ChartObjects('GraphifyChart')
    $s1 = $co.Chart.SeriesCollection(1)
    L "PJ lastNonEmptyI=$($pj.Range('I1048576').End(-4162).Row) points=$($s1.Points.Count) type=$($co.Chart.ChartType) title='$($co.Chart.ChartTitle.Text)'"
    try {
        $v = $s1.Values
        L "PJ series v(first)=$($v[1,1]) I8=$($pj.Range('I8').Value2)"
    } catch { L "values probe skipped: $($_.Exception.Message)" }

    $xl.Run('Graphify')
    L "PJ Graphify run #2 (reuse) OK charts=$($pj.ChartObjects.Count) points=$($pj.ChartObjects('GraphifyChart').Chart.SeriesCollection(1).Points.Count)"

    $gj = $wb.Worksheets('GJ')
    $gj.Activate()
    $gj.Range('A1').Select()
    $xl.Run('Graphify')
    L "GJ Graphify run (auto-pick) OK points=$($gj.ChartObjects('GraphifyChart').Chart.SeriesCollection(1).Points.Count) title='$($gj.ChartObjects('GraphifyChart').Chart.ChartTitle.Text)'"

    $pj.Activate()
    $pj.Range('I10:I20').Select()
    $xl.Run('Graphify')
    L "PJ Graphify run #3 (selection I10:I20) OK points=$($pj.ChartObjects('GraphifyChart').Chart.SeriesCollection(1).Points.Count) (expect 11)"

    $pj.ChartObjects('GraphifyChart').Delete()
    $gj.ChartObjects('GraphifyChart').Delete()
    L "test charts removed"

    # ---------- final CDJ state ----------
    L "CDJ rows 15-16 after fixes:"
    foreach ($r in @(15, 16)) {
        L "  r${r}: F=$($cdj.Range("F$r").Value2) G=$($cdj.Range("G$r").Value2) H=$($cdj.Range("H$r").Value2) N=$($cdj.Range("N$r").Value2) T=$($cdj.Range("T$r").Value2) U=$($cdj.Range("U$r").Value2) Y=$($cdj.Range("Y$r").Text) Z=$($cdj.Range("Z$r").Text)"
    }

    $wb.Save()
    L "saved"
    $wb.Close($false)
    $wb = $null
    L "closed"
    $ok = $true
}
catch {
    L "ERROR: $($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))"
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($xl -ne $null) {
        try { $xl.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    }
    L "excel released; ok=$ok"
    Start-Sleep -Seconds 1
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddMinutes(-30) } |
        ForEach-Object { L "killing leftover EXCEL pid=$($_.Id)"; Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
}
if (-not $watchProc.HasExited) { Stop-Process -Id $watchProc.Id -Force -ErrorAction SilentlyContinue }
if (Test-Path $log) { Get-Content $log }
'--- dialog watch log ---'
if (Test-Path $dlog) { Get-Content $dlog }
if ($ok) { exit 0 } else { exit 1 }
