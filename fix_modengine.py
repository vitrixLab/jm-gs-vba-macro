"""
fix_modengine.py  –  Patch modEngine_v833_extracted.bas for two PJ→CDJ sync bugs

BUG 1 (CRASH): In the INSERT branch of SyncPJToCDJ, line:
   .Cells(targetRow, expenseCol).NumberFormat = "..."
is OUTSIDE the guard "If expenseCol >= CDJ_EXPENSE_START ... Then".
When expenseCol=0 (no COA match), .Cells(row,0) raises Error 1004 which the
empty ErrorHandler silently swallows, killing the whole sync.

FIX 1: Move that NumberFormat line inside the guard block.

BUG 2 (NO MATCH): FindExpenseColumn only uses substring match on the live CDJ
row-13+14 column headers.  Most PJ COA values (e.g. "Clinic Materials and
Supplies", "Rental", "Professional Fees", "Fuel and Oil") do NOT substring-match
the actual CDJ header concatenations.

FIX 2: After the live-header scan fails, fall back to a hard-coded COA alias
table that maps canonical PJ COA strings to their correct CDJ column numbers
(cols 14-19) and returns CDJ_COL_SUNDRY (19) for unmapped COAs so they are
always written somewhere (instead of silently dropped).
"""

import re

INPUT  = "modEngine_v833_extracted.bas"
OUTPUT = "modEngine_v833_fixed.bas"

with open(INPUT, "rb") as f:
    raw = f.read()

# Detect line ending used (file uses \r\r\n from oletools extraction)
LE = b"\r\r\n"

text = raw.decode("utf-8")

# ──────────────────────────────────────────────────────────────────────────────
# FIX 1  –  guard the NumberFormat line for expenseCol in the INSERT branch
# ──────────────────────────────────────────────────────────────────────────────
OLD_INSERT = (
    '            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\r\r\n'
    '                .Cells(targetRow, expenseCol).Value = netAmount\r\r\n'
    '                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount\r\r\n'
    '            End If\r\r\n'
    '            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now\r\r\n'
    '            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '        End With'
)

NEW_INSERT = (
    '            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\r\r\n'
    '                .Cells(targetRow, expenseCol).Value = netAmount\r\r\n'
    '                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount\r\r\n'
    '                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            End If\r\r\n'
    '            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now\r\r\n'
    '            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"\r\r\n'
    '        End With'
)

if OLD_INSERT in text:
    text = text.replace(OLD_INSERT, NEW_INSERT, 1)
    print("[FIX 1] INSERT branch NumberFormat guard applied.")
else:
    print("[FIX 1] WARNING: Target text not found – manual review needed.")
    # Print first differing position for debugging
    for i, (a, b) in enumerate(zip(OLD_INSERT, text[text.find('If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\r\r\n                .Cells(targetRow, expenseCol).Value = netAmount'):text.find('If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\r\r\n                .Cells(targetRow, expenseCol).Value = netAmount')+len(OLD_INSERT)])):
        if a != b:
            print(f"  First diff at offset {i}: expected {repr(a)}, got {repr(b)}")
            break

# ──────────────────────────────────────────────────────────────────────────────
# FIX 2  –  extend FindExpenseColumn with COA alias table fallback
# ──────────────────────────────────────────────────────────────────────────────
OLD_FEC = (
    'Public Function FindExpenseColumn(description As String) As Integer\r\r\n'
    '    Dim descLower As String, i As Integer\r\r\n'
    '    If Not gInitialized Then Call InitializeEngine(gSyncSheet)\r\r\n'
    '    descLower = LCase(Trim(description))\r\r\n'
    '    If descLower = "" Then FindExpenseColumn = 0: Exit Function\r\r\n'
    '    For i = 1 To gHeaderCount\r\r\n'
    '        If arrHeaders(i) <> "" Then\r\r\n'
    '            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then\r\r\n'
    '                FindExpenseColumn = arrCols(i)\r\r\n'
    '                Exit Function\r\r\n'
    '            End If\r\r\n'
    '        End If\r\r\n'
    '    Next i\r\r\n'
    '    FindExpenseColumn = 0\r\r\n'
    'End Function'
)

NEW_FEC = (
    'Public Function FindExpenseColumn(description As String) As Integer\r\r\n'
    '    Dim descLower As String, i As Integer\r\r\n'
    '    If Not gInitialized Then Call InitializeEngine(gSyncSheet)\r\r\n'
    '    descLower = LCase(Trim(description))\r\r\n'
    '    If descLower = "" Then FindExpenseColumn = 0: Exit Function\r\r\n'
    '    \' --- Pass 1: live CDJ header substring match ---\r\r\n'
    '    For i = 1 To gHeaderCount\r\r\n'
    '        If arrHeaders(i) <> "" Then\r\r\n'
    '            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then\r\r\n'
    '                FindExpenseColumn = arrCols(i)\r\r\n'
    '                Exit Function\r\r\n'
    '            End If\r\r\n'
    '        End If\r\r\n'
    '    Next i\r\r\n'
    '    \' --- Pass 2: COA alias table (PJ COA values -> CDJ column) ---\r\r\n'
    '    \' CDJ column map: 14=Clinic Materials, 15=Rent, 16=Gas/Fuel, 17=Comm/Light/Water, 18=Prof Fees, 19=Sundry\r\r\n'
    '    Select Case True\r\r\n'
    '        Case InStr(descLower, "clinic material") > 0 Or InStr(descLower, "clinic supplies") > 0 _\r\r\n'
    '          Or InStr(descLower, "pantry") > 0 Or InStr(descLower, "medical supplies") > 0\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_START       \' col 14\r\r\n'
    '        Case InStr(descLower, "rent") > 0 Or InStr(descLower, "rental") > 0 _\r\r\n'
    '          Or InStr(descLower, "lease") > 0\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_START + 1   \' col 15\r\r\n'
    '        Case InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 _\r\r\n'
    '          Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 _\r\r\n'
    '          Or InStr(descLower, "transport") > 0 Or InStr(descLower, "delivery") > 0\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_START + 2   \' col 16\r\r\n'
    '        Case InStr(descLower, "communication") > 0 Or InStr(descLower, "light") > 0 _\r\r\n'
    '          Or InStr(descLower, "water") > 0 Or InStr(descLower, "utilities") > 0 _\r\r\n'
    '          Or InStr(descLower, "electric") > 0 Or InStr(descLower, "internet") > 0\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_START + 3   \' col 17\r\r\n'
    '        Case InStr(descLower, "professional") > 0 Or InStr(descLower, "consulting") > 0 _\r\r\n'
    '          Or InStr(descLower, "clinician") > 0 Or InStr(descLower, "legal") > 0 _\r\r\n'
    '          Or InStr(descLower, "accounting") > 0\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_START + 4   \' col 18\r\r\n'
    '        Case Else\r\r\n'
    '            \' Unmapped COA -> Sundry column (col 19) so entry is always written\r\r\n'
    '            FindExpenseColumn = CDJ_EXPENSE_END         \' col 19 = Sundry\r\r\n'
    '    End Select\r\r\n'
    'End Function'
)

if OLD_FEC in text:
    text = text.replace(OLD_FEC, NEW_FEC, 1)
    print("[FIX 2] FindExpenseColumn COA alias table applied.")
else:
    print("[FIX 2] WARNING: FindExpenseColumn target text not found.")
    # Try to find partial
    probe = 'Public Function FindExpenseColumn(description As String) As Integer'
    idx = text.find(probe)
    print(f"  Function found at offset: {idx}")

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(text)

print(f"Output written to: {OUTPUT}")
