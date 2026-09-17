"""Apply CDJ plan-vs-implementation fixes to bas files (whitespace-safe)."""
import re

P = r"C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas"
s = open(P, encoding="utf-8").read()
orig = s
n = [0]
def rep(old, new, tag):
    global s
    if old in s:
        s = s.replace(old, new, 1)
        n[0] += 1
        print(f"ok   {tag}")
    else:
        print(f"MISS {tag}")

# 1. Named PNV ref column (Z=26) instead of magic +1
rep("Private Const CDJ_COL_REF As Integer = 25          ' Y (PJ ref) / Z=26 PNV ref",
    "Private Const CDJ_COL_REF As Integer = 25          ' Y (PJ ref)\n\nPrivate Const CDJ_COL_REF_PNV As Integer = 26      ' Z (PNV ref)",
    "named PNV ref col")

# 2. Headers: combine two-row titles (r13 + r14) so Fuel->P, Rent->O, Comm->Q
rep("    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n        header = Trim(ws.Cells(14, col).value)\n\n        If header <> \"\" Then count = count + 1\n\n    Next col",
    "    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n        header = Trim(Trim(ws.Cells(13, col).value) & \" \" & Trim(ws.Cells(14, col).value))\n\n        If header <> \"\" Then count = count + 1\n\n    Next col",
    "headers count uses r13+r14")
rep("    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n        header = Trim(ws.Cells(14, col).value)\n\n        If header <> \"\" Then",
    "    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n        header = Trim(Trim(ws.Cells(13, col).value) & \" \" & Trim(ws.Cells(14, col).value))\n\n        If header <> \"\" Then",
    "headers store uses r13+r14")

# 3. LoadCDJReferences: scan BOTH Y and Z
rep("""    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

    For i = CDJ_DATA_START To lastRow

        ref = Trim(ws.Cells(i, CDJ_COL_REF).value)

        If ref <> \"\" Then dictRef(ref) = i

    Next i""",
    """    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).row

    For i = CDJ_DATA_START To lastRow

        ref = Trim(ws.Cells(i, CDJ_COL_REF).value)

        If ref <> \"\" Then dictRef(ref) = i

        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).value)

        If ref <> \"\" Then dictRef(ref) = i

    Next i""",
    "refs load Y+Z")

# 4. INSERT fallback loops: check both Y and Z
old_loop = """            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).value <> \"\"

                targetRow = targetRow + 1

            Loop"""
new_loop = """            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).value <> \"\" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).value <> \"\"

                targetRow = targetRow + 1

            Loop"""
c = s.count(old_loop)
s = s.replace(old_loop, new_loop)
n[0] += c
print(f"ok   insert loops check Y+Z x{c}")

# 5. NonVat UPDATE: named PNV col + day D
rep(".Cells(targetRow, CDJ_COL_REF + 1).value = ref\n\n            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n                .Cells(targetRow, c).value = \"\"\n\n            Next c\n\n            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated Non-Vat Row",
    ".Cells(targetRow, CDJ_COL_REF_PNV).value = ref\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END\n\n                .Cells(targetRow, c).value = \"\"\n\n            Next c\n\n            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Updated Non-Vat Row",
    "NonVat UPDATE PNV col + day + debit")

# 6. NonVat INSERT: named PNV col + day D + debit=net
rep(".Cells(targetRow, CDJ_COL_REF + 1).value = ref\n\n            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted Non-Vat Row",
    ".Cells(targetRow, CDJ_COL_REF_PNV).value = ref\n\n            .Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value\n\n            .Cells(targetRow, expenseCol).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount\n\n            .Cells(targetRow, CDJ_COL_NOTE).value = \"Inserted Non-Vat Row",
    "NonVat INSERT PNV col + day + debit")

if s != orig:
    open(P, "w", encoding="utf-8").write(s)
print(f"modEngine.bas: {n[0]} fixes applied")

# --- Sheet9: dead finder searched col X; point at Y:Z ---
P9 = r"C:\citrixlabph\globalsmile\_vba_extract\bas\Sheet9.cls"
s9 = open(P9, encoding="utf-8").read()
old9 = "Set found = wsCDJ.Columns(24).Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)"
new9 = ("Set found = wsCDJ.Range(\"Y:Y\").Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)\n\n"
        "    If found Is Nothing Then Set found = wsCDJ.Range(\"Z:Z\").Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)")
if old9 in s9:
    s9 = s9.replace(old9, new9, 1)
    open(P9, "w", encoding="utf-8").write(s9)
    print("ok   Sheet9 finder searches Y:Z")
else:
    print("MISS Sheet9 finder")
