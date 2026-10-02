$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $false

$wb = $xl.Workbooks.Open("D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm", 0, $false)
$name = $wb.Name

Write-Host "WB Name: $name"

# Test syntax 1: just macro name
try {
    $xl.Run("InitializeEngine", "CDJ")
    Write-Host "Syntax 1 (InitializeEngine) -> SUCCESS!"
} catch {
    Write-Host "Syntax 1 failed: $($_.Exception.Message)"
}

# Test syntax 2: $name!InitializeEngine
try {
    $xl.Run("$name!InitializeEngine", "CDJ")
    Write-Host "Syntax 2 ($name!InitializeEngine) -> SUCCESS!"
} catch {
    Write-Host "Syntax 2 failed: $($_.Exception.Message)"
}

# Test syntax 3: RefreshGL
try {
    $xl.Run("RefreshGL")
    Write-Host "RefreshGL -> SUCCESS!"
} catch {
    Write-Host "RefreshGL failed: $($_.Exception.Message)"
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
