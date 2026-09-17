"""Contrast modEngine.bas vs Global-Smile_2026-v7.8_vba.txt for CDJ writer path.

In v7.8 global-vba, the PJ/CDJ writer (SyncPJToCDJ) writes:
    debit = net + whtax   (col T)
    category header T => uses CDJ Two-tier header 13/14
    Exp cols = CDJ cols 14..19
    No Shield N/S => uses CDJ "N" (Shield)?

In modEngine.bas (7.9-patched basis), the PJ/CDJ writer uses:
    debit = netAmount + whtax
    expenseCol from CDJ headers 13/14
    col T guard: writes netAmount into T only if expenseCol in 14..18,
        else if expenseCol == 19 then T gets cat+netWHT? (need to read)
        else T empty?
    In T=0 rows, no debit written.
    non-Vat path: same logic.

The key constant: CDJ_EXPENSE_START=14, CDJ_EXPENSE_END=19.
In v7.8, CDJ old had rows 15..67 filled, each column 14..19 header cells
read; CDJ (patch) uses same range but has real values written into col T only
when category matches.

Goal: ensure the modEngine written values match what CDJ old had in col T,
and category headers are two-tier (13+14).
"""
import re
s = open(P, encoding="utf-8").read()
old = """    expenseCol = FindExpenseColumn(description)

    whtax = CalculateWithholdingTax(wsTarget.Cells(14, expenseCol).value, grossAmount)"""

new = """    Set wsCDJ = ThisWorkbook.Sheets(gSyncSheet)

    expenseCol = FindExpenseColumn(description)

    If expenseCol = 0 Or (expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END) Then

        whtax = 0

    Else

        whtax = CalculateWithholdingTax(Trim(Trim(wsCDJ.Cells(13, expenseCol).value) & " " & Trim(wsCDJ.Cells(14, expenseCol).value)), grossAmount)

    End If"""

assert old in P and old in open(P, encoding="utf-8").read()
new = """    expenseCol = FindExpenseColumn(description)

    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

        whtax = CalculateWithholdingTax(Trim(Trim(wsTarget.Cells(13, expenseCol).value) & \" \" & Trim(wsTarget.Cells(14, expenseCol).value)), grossAmount)

    Else

        whtax = 0

    End If"""
assert old in s, "PJ WHT block not found"
s = s.replace(old, new, 1)
open(P, "w", encoding="utf-8").write(s)
print("ok WHT guarded + uses combined r13+r14 header")
