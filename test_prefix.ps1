$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
$wbName = $wb.Name

Write-Host "WB Name: $wbName"
try {
    $res1 = $xl.Run("'$wbName'!RefreshGL")
    Write-Host "Direct Run('$wbName'!RefreshGL): SUCCESS"
} catch {
    Write-Host "Direct Run('$wbName'!RefreshGL) failed: $($_.Exception.Message)"
}

try {
    $res2 = $xl.Run("'$wbName'!FindExpenseColumn", "Clinic Materials and Supplies")
    Write-Host "Direct Run('$wbName'!FindExpenseColumn): $res2"
} catch {
    Write-Host "Direct Run('$wbName'!FindExpenseColumn) failed: $($_.Exception.Message)"
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
