"""Fix Sheet9.cls compilation error in the .xlsm workbook.

The issue: Sheet9's Worksheet_Change event calls CalculateTotals()
but the CalculateTotals subroutine was not defined in Sheet9's module
(only in Sheet11/Sheet12). This causes a "Sub or Function not defined"
compilation error.

This script:
1. Extracts the VBA code from the .xlsm
2. Adds the missing CalculateTotals and FindBlockStart subroutines
3. Repacks the .xlsm
"""
import zipfile
import shutil
import os

XLsm_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.1.xlsm"
BACKUP_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.1_backup.xlsm"
PATCHED_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.1_patched.xlsm"

# Step 1: Copy the original to backup
shutil.copy2(XLsm_PATH, BACKUP_PATH)
print(f" backed up to {BACKUP_PATH}")

# Step 2: Extract the vbaProject.bin from the ZIP
with zipfile.ZipFile(BACKUP_PATH, 'r') as zin:
    vba_data = zin.read("xl/vbaProject.bin")

print(f" VBA project extracted: {len(vba_data)} bytes")

# Step 3: Check if CalculateTotals already exists in the VBA
vba_text = vba_data.decode("utf-8", errors="replace")
has_calculatetotals = "CalculateTotals" in vba_text
has_findblockstart = "FindBlockStart" in vba_text

print(f" CalculateTotals already present: {has_calculatetotals}")
print(f" FindBlockStart already present: {has_findblockstart}")

if has_calculatetotals and has_findblockstart:
    print(" Sheet9 already has the required subroutines. Nothing to fix.")
    exit(0)

# Step 4: Add the missing subroutines
# We need to insert CalculateTotals and FindBlockStart into Sheet9's module
# The Sheet9 module starts with "Attribute VB_Name = \"Sheet9\""

# Find the Sheet9 module block
# VBA modules are separated by the Attribute VB_Name lines
modules = []
current_module = []
in_sheet9 = False

for line in vba_text.split("\n"):
    if "Attribute VB_Name" in line:
        # Extract the module name
        if "Sheet9" in line:
            in_sheet9 = True
        else:
            in_sheet9 = False
        # Save previous module if any
        if current_module and not in_sheet9:
            modules.append("\n".join(current_module))
        current_module = [line]
    elif in_sheet9:
        current_module.append(line)
    else:
        # Non-Sheet9 module, collect outside
        pass

# Don't forget the last module
if current_module:
    modules.append("\n".join(current_module))

print(f" Found {len(modules)} VBA modules")

# Find Sheet9 module
sheet9_module = None
for mod in modules:
    if "Attribute VB_Name = \"Sheet9\"" in mod or "Attribute VB_Name = 'Sheet9'" in mod:
        sheet9_module = mod
        break

if sheet9_module is None:
    print("ERROR: Could not find Sheet9 module!")
    exit(1)

print(f" Sheet9 module found: {len(sheet9_module)} bytes")

# Add the CalculateTotals and FindBlockStart at the end of Sheet9 module
# (before the next module or end of file)
calculatetotals_sub = '''\nPrivate Sub CalculateTotals(rowNum As Long)\n\n    Dim startRow As Long, endRow As Long, debitSum As Double, creditSum As Double\n\n    startRow = FindBlockStart(rowNum): endRow = rowNum - 1\n\n    If endRow < startRow Then Exit Sub\n\n    debitSum = 0#: creditSum = 0#\n\n    Dim r As Long\n\n    For r = startRow To endRow\n\n        If IsNumeric(Me.Cells(r, COL_DEBIT).Value) Then debitSum = debitSum + Me.Cells(r, COL_DEBIT).Value\n\n        If IsNumeric(Me.Cells(r, COL_DEBIT_MIRROR).Value) Then creditSum = creditSum + Me.Cells(r, COL_DEBIT_MIRROR).Value\n\n    Next r\n\n    Me.Cells(rowNum, COL_DEBIT).Value = debitSum\n\n    Me.Cells(rowNum, COL_DEBIT_MIRROR).Value = debitSum\n\n    Me.Cells(rowNum, 13).Value = creditSum\n\n    Me.Cells(rowNum, 12).Value = creditSum\n\nEnd Sub\n\nPrivate Function FindBlockStart(rowNum As Long) As Long\n\n    Dim i As Long\n\n    For i = rowNum - 1 To START_ROW Step -1\n\n        If Me.Cells(i, COL_COA).Value <> "" Then\n\n            FindBlockStart = i + 1\n\n            Exit Function\n\n        End If\n\n    Next i\n\n    FindBlockStart = START_ROW\n\nEnd Function'''

# Insert the subroutines before the next Attribute VB_Name or end
if "CalculateTotals" not in sheet9_module:
    # Insert before the end of module
    sheet9_module_updated = sheet9_module.rstrip() + calculatetotals_sub
    # Replace in the full VBA text
    vba_text = vba_text.replace(sheet9_module, sheet9_module_updated)
    print(" Added CalculateTotals and FindBlockStart to Sheet9")

# Step 5: Repack the .xlsm
with zipfile.ZipFile(PATCHED_PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.namelist():
        if item == "xl/vbaProject.bin":
            zout.writet(item, vba_text.encode("utf-8"))
        else:
            zout.writet(item, zin.read(item))

print(f" Patched workbook written to {PATCHED_PATH}")

# Verify
with zipfile.ZipFile(PATCHED_PATH, 'r') as zver:
    verify_vba = zver.read("xl/vbaProject.bin").decode("utf-8", errors="replace")
    verify_has_ct = "CalculateTotals" in verify_vba
    verify_has_fbs = "FindBlockStart" in verify_vba
    print(f" Verification - CalculateTotals: {verify_has_ct}")
    print(f" Verification - FindBlockStart: {verify_has_fbs}")
"