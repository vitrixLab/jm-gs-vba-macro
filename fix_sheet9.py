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
# Find the Sheet9 module and add CalculateTotals + FindBlockStart

# Split VBA into modules by Attribute VB_Name markers
module_starts = []
for i, line in enumerate(vba_text.split("\n")):
    if "Attribute VB_Name" in line:
        module_starts.append(i)

if len(module_starts) < 2:
    print("ERROR: Could not parse VBA modules properly")
    exit(1)

# Get Sheet9 module (should be between first and second VB_Name markers, or after first)
# Actually, let's find Sheet9 specifically
sheet9_start = None
sheet9_end = None

for i, line in enumerate(vba_text.split("\n")):
    if "Sheet9" in line and "Attribute VB_Name" in line:
        sheet9_start = i
        # Find the next Attribute VB_Name after this one
        for j in range(i+1, len(vba_text.split("\n"))):
            if "Attribute VB_Name" in vba_text.split("\n")[j]:
                sheet9_end = j
                break
        break

if sheet9_start is None:
    print("ERROR: Could not find Sheet9 module!")
    exit(1)

# Actually let me just get lines from sheet9_start to the end or next module

all_lines = vba_text.split("\n")
if sheet9_end:
    sheet9_lines = all_lines[sheet9_start:sheet9_end]
else:
    sheet9_lines = all_lines[sheet9_start:]

sheet9_code = "\n".join(sheet9_lines)
print(f" Sheet9 module: {len(sheet9_lines)} lines, {len(sheet9_code)} chars")

# Check if CalculateTotals already in Sheet9
if "CalculateTotals" in sheet9_code:
    print(" CalculateTotals already in Sheet9")
else:
    # Add CalculateTotals and FindBlockStart at the end of Sheet9 module
    new_subs = '''
Private Sub CalculateTotals(rowNum As Long)

    Dim startRow As Long, endRow As Long, debitSum As Double, creditSum As Double

    startRow = FindBlockStart(rowNum): endRow = rowNum - 1

    If endRow < startRow Then Exit Sub

    debitSum = 0#: creditSum = 0#

    Dim r As Long

    For r = startRow To endRow

        If IsNumeric(Me.Cells(r, COL_DEBIT).Value) Then debitSum = debitSum + Me.Cells(r, COL_DEBIT).Value

        If IsNumeric(Me.Cells(r, COL_DEBIT_MIRROR).Value) Then creditSum = creditSum + Me.Cells(r, COL_DEBIT_MIRROR).Value

    Next r

    Me.Cells(rowNum, COL_DEBIT).Value = debitSum

    Me.Cells(rowNum, COL_DEBIT_MIRROR).Value = debitSum

    Me.Cells(rowNum, 13).Value = creditSum

    Me.Cells(rowNum, 12).Value = creditSum

End Sub

Private Function FindBlockStart(rowNum As Long) As Long

    Dim i As Long

    For i = rowNum - 1 To START_ROW Step -1

        If Me.Cells(i, COL_COA).Value <> "" Then

            FindBlockStart = i + 1

            Exit Function

        End If

    Next i

    FindBlockStart = START_ROW

End Function'''

    # Insert before the last line if it's Attribute VB_Name, or at end
    if sheet9_lines and "Attribute VB_Name" in sheet9_lines[-1]:
        # Insert before the VB_Name attribute of the next module
        sheet9_code_with_subs = "\n".join(sheet9_lines[:-1]) + new_subs + "\n" + sheet9_lines[-1]
    else:
        sheet9_code_with_subs = sheet9_code + new_subs

    # Replace in full vba_text
    # Find the exact position and replace
    before = "\n".join(all_lines[:sheet9_start])
    after = "\n".join(all_lines[sheet9_end if sheet9_end else len(all_lines):])
    new_vba_text = before + sheet9_code_with_subs + "\n" + after

    vba_text = new_vba_text
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