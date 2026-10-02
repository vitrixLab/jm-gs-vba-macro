$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
$name = $wb.Name

Write-Host "Trying '$name'!modEngine.InitializeEngine ..."
try {
    $xl.Run("'$name'!modEngine.InitializeEngine", "CDJ")
    Write-Host "Success 1"
} catch {
    Write-Host "Failed 1: $($_.Exception.Message)"
}

Write-Host "Trying '$name'!FindExpenseColumn ..."
try {
    $res = $xl.Run("'$name'!FindExpenseColumn", "Clinic Materials and Supplies")
    Write-Host "Success 2: result=$res"
} catch {
    Write-Host "Failed 2: $($_.Exception.Message)"
}

Write-Host "Trying RefreshGL ..."
try {
    $xl.Run("'$name'!RefreshGL")
    Write-Host "Success 3"
} catch {
    Write-Host "Failed 3: $($_.Exception.Message)"
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
