"""Apply CDJ alignment fixes to Module1.bas and Sheet9.cls (engine delegation)."""
import re
import sys

BASE = r"C:\citrixlabph\globalsmile\_vba_extract\bas"


def load(name):
    with open(f"{BASE}\\{name}", encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


def save(name, text):
    with open(f"{BASE}\\{name}", "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)


def replace_sub(text, sub_head_regex, new_body, tag):
    m = re.search(sub_head_regex, text, re.M)
    if not m:
        print(f"FAIL {tag}: sub not found")
        sys.exit(1)
    end = text.index("\nEnd Sub", m.end())
    print(f"ok   {tag}")
    return text[: m.start()] + new_body.rstrip("\n") + "\n" + text[end + len("\nEnd Sub"):]


WRAPPER = """Sub CreateCDJFromPurchase(ByVal sourceSheet As Worksheet, ByVal rowNum As Long)

    ' Delegates to the engine writers so CDJ values land where the titles say:
    ' C=Date, E=Particulars, F=Cash In Bank, G=Input VAT, H=WHT, N-S=category,
    ' T=Debit (net + WHT), X=note, Y/Z=reference, AA13/AA14=sequence + row counter.

    If UCase$(sourceSheet.Name) = "PJ NON-VAT" Then

        modEngine.SyncPJNonVatToCDJ sourceSheet, rowNum

    Else

        modEngine.SyncPJToCDJ sourceSheet, rowNum

    End If

End Sub
"""

BACKFILL = """Sub BackfillAllAutomation()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim ref As String

    ' --- Process PJ rows (vatable) ---

    Set ws = ThisWorkbook.Sheets("PJ")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For i = 8 To lastRow

        ref = Trim(ws.Cells(i, "H").value)

        If ref <> "" And UCase$(ref) <> "TOTAL" Then

            If modEngine.FindCDJRow(ref) = 0 Then

                CreateCDJFromPurchase ws, i

            End If

        End If

    Next i

    ' --- Process PJ Non-Vat rows ---

    Set ws = ThisWorkbook.Sheets("PJ Non-Vat")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For i = 8 To lastRow

        ref = Trim(ws.Cells(i, "H").value)

        If ref <> "" And UCase$(ref) <> "TOTAL" Then

            If modEngine.FindCDJRow(ref) = 0 Then

                CreateCDJFromPurchase ws, i

            End If

        End If

    Next i

    MsgBox "Backfill complete! Processed all existing rows.", vbInformation

End Sub
"""

DELEGATION = """Private Sub SyncToCDJ_NonVat(sourceRow As Long)

    ' CDJ values must land where the titles say (C=Date, E=Particulars, F=Cash In
    ' Bank, G=Input VAT, H=WHT, N-S=category, T=Debit, X/Y/Z=logs) - the engine
    ' owns that layout now, so delegate instead of writing A-E/M-T here.

    modEngine.SyncPJNonVatToCDJ Me, sourceRow

End Sub
"""

t = load("Module1.bas")
t = replace_sub(t, r"^Sub CreateCDJFromPurchase\(ByVal sourceSheet As Worksheet, ByVal rowNum As Long\)\s*$", WRAPPER,
                "Module1 CreateCDJFromPurchase -> engine wrapper")
t = replace_sub(t, r"^Sub BackfillAllAutomation\(\)\s*$", BACKFILL,
                "Module1 BackfillAllAutomation -> ref-based dedup")
save("Module1.bas", t)

t = load("Sheet9.cls")
t = replace_sub(t, r"^Private Sub SyncToCDJ_NonVat\(sourceRow As Long\)\s*$", DELEGATION,
                "Sheet9 SyncToCDJ_NonVat -> engine delegation")
save("Sheet9.cls", t)

print("Module1.bas + Sheet9.cls: ALL FIXES APPLIED")
