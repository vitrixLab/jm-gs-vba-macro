$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $false

$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
$name = $wb.Name

Write-Host "Workbook opened."
Write-Host "Running InitializeEngine via COM..."
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$xl.Run("'$name'!modEngine.InitializeEngine", "CDJ")
$sw.Stop()
Write-Host "InitializeEngine finished in $($sw.Elapsed.TotalSeconds)s!"

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
