$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $false

Write-Host "Opening workbook with EnableEvents = false..."
$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
Write-Host "Opened successfully!"
$name = $wb.Name
Write-Host "Workbook Name: $name"

# Let's inspect ThisWorkbook.cls in the open workbook
$compTW = $wb.VBProject.VBComponents.Item("ThisWorkbook")
$twCode = $compTW.CodeModule.Lines(1, $compTW.CodeModule.CountOfLines)
Write-Host "ThisWorkbook code:`n$twCode"

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
