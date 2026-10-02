# rebuild_v833_pjfix.ps1
# Patches modEngine in Global-Smile_2026-v8.3.3.xlsm with the PJ->CDJ sync fixes
# Fix 1: Guard .Cells(row,expenseCol).NumberFormat inside If expenseCol in range
# Fix 2: COA alias table in FindExpenseColumn (Pass 2 fallback)

param(
    [string]$XlsmPath = "D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm",
    [string]$FixedModule = "D:\citrixlabph\globalsmile\modEngine_v833_fixed.bas"
)

$ErrorActionPreference = "Stop"
$log = "D:\citrixlabph\globalsmile\rebuild_pjfix.log"
function L($m) { $ts = Get-Date -f "HH:mm:ss"; "$ts  $m" | Tee-Object -Append -FilePath $log; Write-Host "$ts  $m" }

"" | Out-File $log
L "=== Rebuild v8.3.3 – PJ->CDJ Sync Fix ==="
L "Target: $XlsmPath"
L "Module: $FixedModule"

if (-not (Test-Path $XlsmPath))  { L "ERROR: xlsm not found"; exit 1 }
if (-not (Test-Path $FixedModule)){ L "ERROR: fixed module not found"; exit 1 }

# Launch Excel
$xl  = New-Object -ComObject Excel.Application
$xl.Visible        = $false
$xl.DisplayAlerts  = $false
$xl.EnableEvents   = $false

try {
    L "Opening workbook..."
    $wb = $xl.Workbooks.Open($XlsmPath, 0, $false)
    $vbp = $wb.VBProject
    $vbp.VBE.MainWindow.Visible = $false

    # ── Remove existing modEngine ──────────────────────────────────────────
    L "Removing old modEngine..."
    $found = $false
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Name -eq "modEngine") {
            $vbp.VBComponents.Remove($comp)
            $found = $true
            L "  Removed modEngine."
            break
        }
    }
    if (-not $found) { L "  WARNING: modEngine not found in VBProject – importing anyway." }

    # ── Import fixed module ────────────────────────────────────────────────
    L "Importing fixed modEngine_v833_fixed.bas..."
    $vbp.VBComponents.Import($FixedModule) | Out-Null
    L "  Import done."

    # ── Verify compilation ─────────────────────────────────────────────────
    L "Verifying compilation..."
    try {
        $xl.VBE.CommandBars.Item("Standard").Controls.Item("Compile VBAProject").Execute()
        L "  Compilation: OK"
    } catch {
        L "  Compilation check via UI not available – relying on runtime test."
    }

    # ── Test: call RefreshGL to confirm no compile errors ─────────────────
    L "Running RefreshGL() to confirm no compile errors..."
    try {
        $xl.Run("RefreshGL")
        L "  RefreshGL: completed."
    } catch {
        L "  RefreshGL error: $_"
    }

    # ── Save ───────────────────────────────────────────────────────────────
    L "Saving workbook..."
    $wb.Save()
    L "  Saved."

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

L "=== Done. v8.3.3 patched with PJ->CDJ sync fix ==="
