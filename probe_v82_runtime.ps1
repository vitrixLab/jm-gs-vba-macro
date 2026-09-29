# probe_v82_runtime.ps1 - READ-ONLY runtime probe of Global-Smile_2026-v8.2-final-review.xlsm
# Opens the workbook read-only with events suppressed, calls only non-interactive
# VBA functions (ValidateGLConsistency, CalculateMonthlyNet), never saves.
$ErrorActionPreference = 'Stop'
$path = 'D:\citrixlabph\globalsmile\Global-Smile_2026-v8.2-final-review.xlsm'
$hashBefore = (Get-FileHash $path -Algorithm SHA256).Hash
Write-Output "workbook sha256 before: $hashBefore"

$accounts = @(
 'Cash in Bank','Petty Cash Fund','Account Receivables','Advances to Employees',
 'Input VAT','Excess CWT Over IT','Leasehold Improvements','Dental Equipment',
 'Due To Clinicians','Due to Clinicians- Visiting','VAT Payable','EWT Payable',
 'Government Contributions (EE)','Government Loans (EE)','Income Tax Payable',
 'Paid-Up Capital','Retained Earnings (Deficit)','Sales','Rent',
 "Clinician's Fee-Corporators","Clinician's Fee-Visiting",'Clinic Material and Supplies',
 'Common Use Service Area','Light and Water Expense',
 "Dep'n Expense-Leasehold Improvements","Dep'n Expense-Dental Equipment",
 'Salaries and Wages','De Minimis','Government Contributions ER Share',
 'Professional Fees','Repair and Maintenance','Pantry Supplies','Office Supplies',
 'Meals, Gifts & Entertainment Expense','Communication ','Printing and Duplication',
 'Taxes and Licences','Delivery','Software Subscription','Miscellaneous',
 'Transportation and Travel','Advertisement','Marketing Expense',
 'Gas, Oil, Parking, Toll Fees','Other Penalties and Charges','Insurance')

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow (allow macros to run)

$wb = $null
try {
    $wb = $excel.Workbooks.Open($path, 0, $true)
    Write-Output "opened read-only: $($wb.Name)  sheet count=$($wb.Worksheets.Count)"

    $ok = $excel.Run("ValidateGLConsistency")
    Write-Output ("ValidateGLConsistency('GL') as shipped  -> " + $ok)

    $total = 0.0
    $nonZero = 0
    $zeroAccounts = New-Object System.Collections.ArrayList
    $perAccount = @{}
    foreach ($a in $accounts) {
        $sum = 0.0
        $n = 0
        for ($m = 1; $m -le 12; $m++) {
            $v = [double]$excel.Run("CalculateMonthlyNet", $a, $m, 2026)
            $sum += $v
            if ([math]::Abs($v) -gt 0.0001) { $n++; $nonZero++ }
        }
        $total += $sum
        $perAccount[$a] = @($sum, $n)
        if ($n -eq 0) { [void]$zeroAccounts.Add($a) }
    }
    Write-Output ("SUM of all 552 CalculateMonthlyNet(debit-credit) values = " + ('{0:N2}' -f $total))
    Write-Output ("account-month cells with activity = $nonZero of 552")
    Write-Output ("accounts with zero activity in all 12 months = " + $zeroAccounts.Count)
    Write-Output ("  " + ($zeroAccounts -join '; '))
    Write-Output "accounts with activity (account : net for year : months with activity):"
    foreach ($a in $accounts) {
        $x = $perAccount[$a]
        if ($x[1] -gt 0) { Write-Output ("   {0,-40} net={1,12:N2}  months={2}" -f $a, $x[0], $x[1]) }
    }
    Write-Output ("Final ValidateGLConsistency() after probes -> " + $excel.Run("ValidateGLConsistency"))
}
finally {
    if ($wb -ne $null) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
$hashAfter = (Get-FileHash $path -Algorithm SHA256).Hash
Write-Output "workbook sha256 after : $hashAfter"
Write-Output ("source unchanged: " + ($hashBefore -eq $hashAfter))
Get-ChildItem 'D:\citrixlabph\globalsmile' -Filter '~$*.xlsm' -Force | Select-Object Name
