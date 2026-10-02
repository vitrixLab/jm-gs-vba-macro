$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $false

$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
$name = $wb.Name

Write-Host "Testing individual loader calls..."

# We can test by running a small VBA sub or inspecting what each loader does
# Let's inspect the sheets directly:
$wsCDJ = $wb.Sheets.Item("CDJ")
$wsSup = $wb.Sheets.Item("SUPPLIERS DATA")

Write-Host "CDJ UsedRange: $($wsCDJ.UsedRange.Address)"
Write-Host "SUPPLIERS DATA UsedRange: $($wsSup.UsedRange.Address)"

# Check CDJ columns 25 and 26 end up row
$lastRef = $wsCDJ.Cells.Item($wsCDJ.Rows.Count, 25).End(-4162).Row # xlUp = -4162
$lastPnv = $wsCDJ.Cells.Item($wsCDJ.Rows.Count, 26).End(-4162).Row
Write-Host "CDJ Col 25 last row: $lastRef, Col 26 last row: $lastPnv"

# Check SUPPLIERS DATA column 3 end up row
$lastSup = $wsSup.Cells.Item($wsSup.Rows.Count, 3).End(-4162).Row
Write-Host "SUPPLIERS DATA Col 3 last row: $lastSup"

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
