"""Fix remaining CDJ plan-vs-implementation gaps in modEngine.bas."""
P = r"C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas"
s = open(P, encoding="utf-8").read()
n = 0
def rep(old, new, tag):
    global s, n
    c = s.count(old)
    if c:
        s = s.replace(old, new)
        n += c
        print(f"ok   {tag} x{c}")
    else:
        print(f"MISS {tag}")

# A. WHT: plan computes from expense-category (col I-S) + gross, not PJ description.
#    Passed `description` is the PJ COA; map it to the CDJ category cell first via
#    FindExpenseColumn, then derive WHT with the same category rules.
rep("    whtax = CalculateWithholdingTax(description, grossAmount)\n\n    expenseCol = FindExpenseColumn(description)",
    "    expenseCol = FindExpenseColumn(description)\n\n    whtax = CalculateWithholdingTax(wsTarget.Cells(14, expenseCol).value, grossAmount)",
    "PJ path: WHT from CDJ category header")

# B. VAT rows: T (Debit) must be gross (net + WHT). netAmount already lands in the
#    category column; T = net + WHT so T - (category) ties to H per plan pattern.
rep("            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated from PJ Row",
    "            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated from PJ Row",
    "PJ UPDATE: T=net+WHT")
rep("            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted from PJ Row",
    "            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted from PJ Row",
    "PJ INSERT: T=net+WHT")

# C. Expense fallback: Sundry (S=19) only when NO category header matches; a blank
#    PJ description must not silently become Sundry. Leave category + T empty then.
rep('    If descLower = "" Then FindExpenseColumn = 19: Exit Function',
    '    If descLower = "" Then FindExpenseColumn = 0: Exit Function',
    "blank description -> no column")
rep("    Next i\n\n    FindExpenseColumn = 19\n\nEnd Function",
    "    Next i\n\n    FindExpenseColumn = 0\n\nEnd Function",
    "no match -> no column")

# D. Guard writers: only touch category/T cells when a category matched.
rep("            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n                .Cells(targetRow, c).value = \"\"\n\n            Next c\n\n            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated from PJ Row",
    "            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n                .Cells(targetRow, c).value = \"\"\n\n            Next c\n\n            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\n\n                .Cells(targetRow, expenseCol).value = netAmount\n\n                .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            End If\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated from PJ Row",
    "PJ UPDATE: guard category/T")
rep("            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted from PJ Row",
    "            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then\n\n                .Cells(targetRow, expenseCol).value = netAmount\n\n                .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            End If\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted from PJ Row",
    "PJ INSERT: guard category/T")

open(P, "w", encoding="utf-8").write(s)
print(f"total replacements: {n}")
