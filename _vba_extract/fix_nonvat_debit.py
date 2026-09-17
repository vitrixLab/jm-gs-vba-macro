P = r'C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas'
s = open(P, encoding='utf-8').read()
orig = s
# NonVat Debit: netAmount -> netAmount + whtax (2 writers: Updated + Inserted Non-Vat)
old = """.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount

            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated Non-Vat Row"""
new = """.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax

            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated Non-Vat Row"""
c1 = s.count(old); s = s.replace(old, new)
old2 = """.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount

            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted Non-Vat Row"""
new2 = """.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax

            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted Non-Vat Row"""
c2 = s.count(old2); s = s.replace(old2, new2)
open(P, "w", encoding="utf-8").write(s)
print(f"NonVat UPDATE debit fixed x{c1}; NonVat INSERT debit fixed x{c2}; changed={s != orig}")
