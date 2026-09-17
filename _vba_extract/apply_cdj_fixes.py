"""Apply CDJ alignment fixes to modEngine.bas (every pattern must hit exactly once)."""
import sys

BASE = r"C:\citrixlabph\globalsmile\_vba_extract\bas"
DEBIT_LINE = "            .Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax"


def load(name):
    with open(f"{BASE}\\{name}", encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


def save(name, text):
    with open(f"{BASE}\\{name}", "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)


def rep(text, old, new, tag):
    n = text.count(old)
    if n != 1:
        print(f"FAIL {tag}: {n} occurrences")
        sys.exit(1)
    print(f"ok   {tag}")
    return text.replace(old, new)


t = load("modEngine.bas")

t = rep(t,
        "Private Const CDJ_EXPENSE_END As Integer = 20      ' T",
        "Private Const CDJ_EXPENSE_END As Integer = 19      ' S (T is the Debit total column, not an expense)",
        "expense range N-S")

t = rep(t,
        "Private Const CDJ_DATA_START As Integer = 15       ' First data row",
        "Private Const CDJ_DATA_START As Integer = 15       ' First data row\n\n"
        "Private Const CDJ_COL_DEBIT As Integer = 20         ' T (Debit) = net + expanded WHT = gross",
        "CDJ_COL_DEBIT const")

t = rep(t,
        'arrHeaders = Array("Clinic Supplies", "Rent", "Fuel", "Utilities", "Professional", "Sundry", "Supplies")',
        'arrHeaders = Array("Clinic Supplies", "Rent", "Fuel", "Utilities", "Professional", "Sundry")',
        "fallback headers")

t = rep(t,
        "arrCols = Array(14, 15, 16, 17, 18, 19, 20)",
        "arrCols = Array(14, 15, 16, 17, 18, 19)",
        "fallback cols")

t = rep(t,
        '            .Cells(targetRow, expenseCol).value = netAmount\n\n'
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated from PJ Row "',
        "            .Cells(targetRow, expenseCol).value = netAmount\n\n" + DEBIT_LINE + "\n\n"
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated from PJ Row "',
        "SyncPJToCDJ UPDATE writes T")

t = rep(t,
        '            .Cells(targetRow, expenseCol).value = netAmount\n\n'
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted from PJ Row "',
        "            .Cells(targetRow, expenseCol).value = netAmount\n\n" + DEBIT_LINE + "\n\n"
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted from PJ Row "',
        "SyncPJToCDJ INSERT writes T")

t = rep(t,
        "            .Cells(targetRow, CDJ_COL_CASH).value = netAmount\n",
        "            .Cells(targetRow, CDJ_COL_CASH).value = netAmount * -1\n",
        "SyncPJNonVatToCDJ UPDATE cash sign fix")

t = rep(t,
        '            .Cells(targetRow, expenseCol).value = netAmount\n\n'
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated Non-Vat Row "',
        "            .Cells(targetRow, expenseCol).value = netAmount\n\n" + DEBIT_LINE + "\n\n"
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated Non-Vat Row "',
        "SyncPJNonVatToCDJ UPDATE writes T")

t = rep(t,
        '            .Cells(targetRow, expenseCol).value = netAmount\n\n'
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted Non-Vat Row "',
        "            .Cells(targetRow, expenseCol).value = netAmount\n\n" + DEBIT_LINE + "\n\n"
        '            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted Non-Vat Row "',
        "SyncPJNonVatToCDJ INSERT writes T")

save("modEngine.bas", t)
print("modEngine.bas: ALL FIXES APPLIED")
