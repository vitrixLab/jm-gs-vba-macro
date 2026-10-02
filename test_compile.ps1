$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open('D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm', 0, $false)
try {
    $xl.VBE.CommandBars.Item('Standard').Controls.Item('Compile VBAProject').Execute()
    Write-Host 'VBA Project compiled cleanly with NO errors!'
} catch {
    Write-Host "Compilation error: $($_.Exception.Message)"
}
try {
    $xl.Run("RefreshGL")
    Write-Host "RefreshGL ran successfully!"
} catch {
    Write-Host "RefreshGL error: $($_.Exception.Message)"
}
try {
    $res = $xl.Run("FindExpenseColumn", "Clinic Materials and Supplies")
    Write-Host "FindExpenseColumn returned: $res"
} catch {
    Write-Host "FindExpenseColumn error: $($_.Exception.Message)"
}
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
