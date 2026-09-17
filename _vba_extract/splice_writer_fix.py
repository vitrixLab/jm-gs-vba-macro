"""Splice-fix SyncPJToCDJ UPDATE + INSERT writers (line-based, robust)."""
P = r"C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas"
lines = open(P, encoding="utf-8").read().split("\n")
out = []
i = 0
fixed_update = fixed_insert = fixed_whtax = 0
while i < len(lines):
    L = lines[i]
    # WHT base: grossAmount -> netAmount inside the SyncPJToCDJ guard
    if "whtax = CalculateWithholdingTax(Trim(Trim(wsTarget.Cells(13, expenseCol).value)" in L and "grossAmount)" in L:
        L = L.replace("grossAmount)", "netAmount)")
        fixed_whtax += 1
        out.append(L); i += 1; continue
    # UPDATE writer: '.Cells(targetRow, expenseCol).value = netAmount' followed by debit line then NOTE; only inside "Updated from PJ Row"
    if L.strip() == ".Cells(targetRow, expenseCol).value = netAmount" and i + 4 < len(lines) \
       and "CDJ_COL_DEBIT" in lines[i+2] and "Updated from PJ Row" in lines[i+4]:
        out.append(L); out.append(lines[i+1])
        out.append(lines[i+2].replace("netAmount + whtax", '""'))
        out.append(lines[i+3]); out.append(lines[i+4]); i += 5
        fixed_update += 1
        continue
    # INSERT writer: same pattern but 'Inserted from PJ Row'
    if L.strip() == ".Cells(targetRow, expenseCol).value = netAmount" and i + 10 < len(lines) \
       and "CDJ_COL_DEBIT" in lines[i+2] and "Inserted from PJ Row" in lines[i+4]:
        out.append(L); out.append(lines[i+1])
        out.append(lines[i+2].replace("netAmount + whtax", '""'))
        for k in range(3, 11):
            out.append(lines[i+k])
        i += 11
        fixed_insert += 1
        continue
    out.append(L); i += 1
open(P, "w", encoding="utf-8").write("\n".join(out))
print(f"whtax base fixed x{fixed_whtax}; UPDATE debit->blank x{fixed_update}; INSERT debit->blank x{fixed_insert}")
