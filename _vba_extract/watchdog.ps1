$ErrorActionPreference = 'Continue'
$wlog = 'C:\citrixlabph\globalsmile\_vba_extract\watchdog.log'
function W($m) { "$(Get-Date -Format 'HH:mm:ss.fff')  $m" | Out-File -FilePath $wlog -Append -Encoding utf8 }
Remove-Item $wlog -ErrorAction SilentlyContinue
W "watchdog start"

# kill stale EXCEL from previous attempts (started in the last 10 min)
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddMinutes(-10) } | ForEach-Object {
    W "killing stale EXCEL pid=$($_.Id)"
    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
}

# stash + enable AccessVBOM
$key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$prop = Get-ItemProperty $key -ErrorAction SilentlyContinue
$had = $false
$old = 0
if ($prop -and $prop.PSObject.Properties['AccessVBOM']) { $had = $true; $old = $prop.AccessVBOM }
W "stash AccessVBOM had=$had old=$old"
if (-not (Test-Path $key)) { New-Item $key -Force | Out-Null }
Set-ItemProperty $key -Name AccessVBOM -Value 1

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$p = Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','C:\citrixlabph\globalsmile\_vba_extract\inject_and_test.ps1' -PassThru -WindowStyle Hidden
$exited = $p.WaitForExit(300000)
if (-not $exited) {
    W "TIMEOUT - killing test + fresh EXCEL"
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
} else {
    W "test exit=$($p.ExitCode)"
}

# restore AccessVBOM
if ($had) { Set-ItemProperty $key -Name AccessVBOM -Value $old } else { Remove-ItemProperty $key -Name AccessVBOM -ErrorAction SilentlyContinue }
W "AccessVBOM restored (had=$had old=$old)"
W "excel processes now: $(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count)"
W "watchdog done"
