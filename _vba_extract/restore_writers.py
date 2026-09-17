P = r'C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas'
s = open(P, encoding='utf-8').read()
orig = s
fixes = []

# A. Restore If-expenseCol split guards to the plain form in the PJ writers
old_if = "If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END - 1 Then"
new_if = "If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then"
c = s.count(old_if)
s = s.replace(old_if, new_if)
fixes.append(f"If-guard restored x{c}")

old_elif = "ElseIf expenseCol = CDJ_EXPENSE_END Then"
c2 = s.count(old_elif)
s = s.replace(old_elif, "End If\n\n            If False Then ' (removed Sundry-description branch)")
fixes.append(f"ElseIf Sundry-description branch neutralized x{c2}")

# B. expenseCol write: description -> netAmount (both S-category branches)
old_desc = ".Cells(targetRow, expenseCol).value = description"
c3 = s.count(old_desc)
s = s.replace(old_desc, ".Cells(targetRow, expenseCol).value = netAmount")
fixes.append(f"expenseCol description->netAmount x{c3}")

# C. Debit: netAmount + whtax -> grossAmount in the PJ writers (lines ~1010, ~1065)
old_dw = ".Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax"
c4 = s.count(old_dw)
s = s.replace(old_dw, ".Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount")
fixes.append(f"debit net+whtax->grossAmount x{c4}")

# D. Day-D: currently NO day writes in PJ writers (only 1921/1972 NonVat) - NOTHING to remove; confirm
open(P, "w", encoding="utf-8").write(s)
print("; ".join(fixes))
print("changed:", s != orig)
