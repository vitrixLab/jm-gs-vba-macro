"""Gate 2 — patch VBA modules for Global-Smile_2026-v8.4-fix.xlsm.

Reads the read-only baseline extraction in gate2_work/vba_baseline/ and writes
patched modules to gate2_work/vba_patched/. The source workbook is not touched.

Every replacement asserts the exact number of occurrences it expects, so a drift
in the baseline fails loudly instead of silently under-patching.
"""
import io
import os
import shutil
import sys

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba_baseline")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba_patched")

NUMFMT = 'NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"'


def patch(name, rules):
    src = os.path.join(BASE, name)
    dst = os.path.join(OUT, name)
    with io.open(src, "r", encoding="utf-8", newline="") as fh:
        text = fh.read()
    nl = "\r\n" if "\r\n" in text else "\n"
    report = []
    for old, new, expected in rules:
        old = old.replace("\n", nl)
        new = new.replace("\n", nl)
        found = text.count(old)
        if found != expected:
            raise SystemExit(
                "PATCH FAIL %s: expected %d occurrence(s), found %d for:\n%s"
                % (name, expected, found, old[:160])
            )
        text = text.replace(old, new)
        report.append("  %dx  %s" % (found, old.strip().split("\n")[0][:78]))
    with io.open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    print("patched %-14s %d rule(s)" % (name, len(rules)))
    for line in report:
        print(line)


# ---------------------------------------------------------------------------
# modEngine.bas
# ---------------------------------------------------------------------------
engine_rules = [
    # R1 - column constants: canonical refs live in AA(27)/AB(28), notes in Z(26).
    (
        """Private Const CDJ_COL_REF As Integer = 25
Private Const CDJ_COL_REF_PNV As Integer = 26
Private Const CDJ_COL_DAY As Integer = 4
Private Const CDJ_COL_NOTE As Integer = 24""",
        """Private Const CDJ_COL_REF As Integer = 27
Private Const CDJ_COL_REF_PNV As Integer = 28
Private Const CDJ_COL_REF_LEGACY As Integer = 25
Private Const CDJ_COL_DAY As Integer = 4
Private Const CDJ_COL_NOTE As Integer = 26""",
        1,
    ),
    # R2 - LoadCDJReferences must also read the legacy Y column.
    (
        """    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
    For i = CDJ_DATA_START To lastRow
        ref = Trim(ws.Cells(i, CDJ_COL_REF).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).Value)
        If ref <> "" Then dictRef(ref) = i
    Next i""",
        """    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
    For i = CDJ_DATA_START To lastRow
        ref = Trim(ws.Cells(i, CDJ_COL_REF).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)
        If ref <> "" And Not dictRef.Exists(ref) Then dictRef(ref) = i
    Next i""",
        1,
    ),
    # R3 - FindCDJRow fallback scan must read the legacy Y column too.
    (
        """        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
        If lastRow >= CDJ_DATA_START Then
            For i = CDJ_DATA_START To lastRow
                If Trim$(CStr(ws.Cells(i, CDJ_COL_REF).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_PNV).Value)) = ref Then""",
        """        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
        If lastRow >= CDJ_DATA_START Then
            For i = CDJ_DATA_START To lastRow
                If Trim$(CStr(ws.Cells(i, CDJ_COL_REF).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_PNV).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)) = ref Then""",
        1,
    ),
    # R4 - stop trusting the static AA14 pointer; resolve a real free row.
    (
        """Private Sub LoadNextRow()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then gNextRow = CDJ_DATA_START: Exit Sub
    val = ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value
    If IsNumeric(val) Then
        gNextRow = CLng(val)
    Else
        gNextRow = CDJ_DATA_START
        Do While ws.Cells(gNextRow, CDJ_COL_REF).Value <> ""
            gNextRow = gNextRow + 1
        Loop
    End If
End Sub""",
        """Private Sub LoadNextRow()
    ' GATE2: the static AA14 pointer is no longer trusted. Excel row insertion
    ' never adjusts it, so it can resolve onto live rows. Resolve the next write
    ' row by scanning for a genuinely empty data row instead (FindCDJFreeRow).
    gNextRow = FindCDJFreeRow(0)
End Sub

Private Function CDJRowIsFree(ws As Worksheet, ByVal r As Long) As Boolean
    If Trim$(CStr(ws.Cells(r, CDJ_COL_PARTICULARS).Value)) <> "" Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_CASH).Value))) > 0 Then Exit Function
    If Trim$(CStr(ws.Cells(r, CDJ_COL_REF).Value)) <> "" Then Exit Function
    If Trim$(CStr(ws.Cells(r, CDJ_COL_REF_PNV).Value)) <> "" Then Exit Function
    If Trim$(CStr(ws.Cells(r, CDJ_COL_REF_LEGACY).Value)) <> "" Then Exit Function
    CDJRowIsFree = True
End Function

Private Function CDJMarkerMonth(ByVal marker As String) As Integer
    Dim t As String
    t = UCase$(Trim$(marker))
    If t = "" Or InStr(t, "TOTAL") > 0 Then Exit Function
    Select Case Left$(t, 3)
        Case "JAN": CDJMarkerMonth = 1
        Case "FEB": CDJMarkerMonth = 2
        Case "MAR": CDJMarkerMonth = 3
        Case "APR": CDJMarkerMonth = 4
        Case "MAY": CDJMarkerMonth = 5
        Case "JUN": CDJMarkerMonth = 6
        Case "JUL": CDJMarkerMonth = 7
        Case "AUG": CDJMarkerMonth = 8
        Case "SEP": CDJMarkerMonth = 9
        Case "OCT": CDJMarkerMonth = 10
        Case "NOV": CDJMarkerMonth = 11
        Case "DEC": CDJMarkerMonth = 12
    End Select
End Function

Private Function CDJMonthAbbr(ByVal m As Integer) As String
    Select Case m
        Case 1: CDJMonthAbbr = "JAN"
        Case 2: CDJMonthAbbr = "FEB"
        Case 3: CDJMonthAbbr = "MAR"
        Case 4: CDJMonthAbbr = "APR"
        Case 5: CDJMonthAbbr = "MAY"
        Case 6: CDJMonthAbbr = "JUN"
        Case 7: CDJMonthAbbr = "JUL"
        Case 8: CDJMonthAbbr = "AUG"
        Case 9: CDJMonthAbbr = "SEP"
        Case 10: CDJMonthAbbr = "OCT"
        Case 11: CDJMonthAbbr = "NOV"
        Case 12: CDJMonthAbbr = "DEC"
    End Select
End Function

Private Function PJMonthOf(sourceSheet As Worksheet, ByVal sourceRow As Long) As Integer
    Dim r As Long, mk As Integer
    For r = 1 To sourceRow
        If UCase$(Trim$(CStr(sourceSheet.Cells(r, 4).Value))) <> "TOTAL" Then
            mk = CDJMarkerMonth(CStr(sourceSheet.Cells(r, 2).Value))
            If mk > 0 Then PJMonthOf = mk
        End If
    Next r
End Function

Public Function FindCDJFreeRow(Optional ByVal targetMonth As Integer = 0) As Long
    Dim ws As Worksheet, r As Long, lastScan As Long, curMonth As Integer, mk As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then FindCDJFreeRow = CDJ_DATA_START: Exit Function
    lastScan = ws.Cells(ws.Rows.Count, CDJ_COL_PARTICULARS).End(xlUp).Row
    If lastScan < CDJ_DATA_START Then lastScan = CDJ_DATA_START
    curMonth = 0
    For r = CDJ_DATA_START To lastScan
        mk = CDJMarkerMonth(CStr(ws.Cells(r, CDJ_COL_DATE).Value))
        If mk > 0 Then curMonth = mk
        If CDJRowIsFree(ws, r) Then
            If targetMonth = 0 Or curMonth = targetMonth Then
                FindCDJFreeRow = r
                Exit Function
            End If
        End If
    Next r
    FindCDJFreeRow = lastScan + 1
End Function""",
        1,
    ),
    # R5 - insert branch: resolve a free row in the target month, write the day.
    (
        """    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then""",
        """    Else
        ' GATE2: place the row inside the target month's own block.
        mth = PJMonthOf(sourceSheet, sourceRow)
        targetRow = FindCDJFreeRow(mth)
        gNextRow = targetRow
        If mth > 0 And Len(Trim$(CStr(dateVal))) = 0 Then dateVal = CDJMonthAbbr(mth)
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then""",
        1,
    ),
    # R6 - update branch: write the day, cash net of WHT, WHT as a credit.
    (
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c""",
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            .Cells(targetRow, CDJ_COL_DEBIT).Value = vbNullString""",
        1,
    ),
    # R7 - expense placement (both branches): T holds sundry net only, never gross.
    (
        """            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: numeric amount into matched column
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry (col 19 = S): COA label text in col S, amount in col R (18)
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
                .Cells(targetRow, CDJ_EXPENSE_END - 1).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If""",
        """            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: net goes into the matched allocation
                ' column only. Column T (Sundry Debit) stays empty, otherwise the
                ' net is counted twice and X no longer balances.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: title in S (19), net in T (20); R (18 = Professional
                ' Fees) stays empty so the amount is not double-classified.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = ""
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If""",
        2,
    ),
    # R8 - declare mth for the insert path (SyncPJToCDJ only).
    (
        """    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))""",
        """    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    Dim mth As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))""",
        1,
    ),
    # R9 - SyncPJNonVatToCDJ update branch: same T double-count, and it wrote
    # the PJ month name into the CDJ day column.
    (
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: numeric amount into matched column
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry (col 19 = S): COA label text in col S, amount in col R (18)
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, CDJ_EXPENSE_END - 1).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now""",
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            .Cells(targetRow, CDJ_COL_DEBIT).Value = ""
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: net into the allocation column only.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: title in S (19), net in T (20); R (18) stays empty.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = ""
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now""",
        1,
    ),
    # R10 - SyncPJNonVatToCDJ insert branch (same treatment).
    (
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: numeric amount into matched column
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry (col 19 = S): COA label text in col S, amount in col R (18)
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, CDJ_EXPENSE_END - 1).""" + NUMFMT + """
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now""",
        """        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18: net into the allocation column only.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).""" + NUMFMT + """
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: title in S (19), net in T (20); R (18) stays empty.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_EXPENSE_END - 1).Value = ""
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).""" + NUMFMT + """
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now""",
        1,
    ),
    # R9b - SyncPJNonVatToCDJ insert branch must not trust the static pointer.
    (
        """    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1""",
        """    Else
        ' GATE2: place the row inside the target month's own block.
        targetRow = FindCDJFreeRow(CDJMarkerMonth(CStr(dateVal)))
        gNextRow = targetRow
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1""",
        1,
    ),
    # R11 - Legacy_* self-references broke compilation under Option Explicit.
    (
        "    CalculateMonthlyNet = x(0) - x(1)",
        "    Legacy_CalculateMonthlyNet = x(0) - x(1)",
        1,
    ),
    (
        """    z(1) = startBalance + CalculateMonthlyNet(accountCode, 1, year)
    For m = 2 To 12
        z(m) = z(m - 1) + CalculateMonthlyNet(accountCode, m, year)
    Next m
    CalculateEndingBalanceFull = z""",
        """    z(1) = startBalance + Legacy_CalculateMonthlyNet(accountCode, 1, year)
    For m = 2 To 12
        z(m) = z(m - 1) + Legacy_CalculateMonthlyNet(accountCode, m, year)
    Next m
    Legacy_CalculateEndingBalanceFull = z""",
        1,
    ),
    (
        "    ValidateGLConsistency = Abs(td - tc) <= GL_AUTO_TOL",
        "    Legacy_ValidateGLConsistency = Abs(td - tc) <= GL_AUTO_TOL",
        1,
    ),
    # R10 - modGLGate calls these two names; provide thin public wrappers.
    (
        """Public Function GetFinancialYear(ByVal dt As Date) As Integer
    GetFinancialYear = year(dt)
End Function""",
        """Public Function GetFinancialYear(ByVal dt As Date) As Integer
    GetFinancialYear = year(dt)
End Function

' ============================================================
' GATE 2: public wrappers required by modGLGate.RunV8Gate.
' modGLGate referenced these names but they did not exist anywhere in the
' project, so the gate could not compile.
' ============================================================
Public Function ValidateGLConsistency(Optional ByVal targetYear As Long = 0) As Boolean
    ValidateGLConsistency = Legacy_ValidateGLConsistency("GL")
End Function

Public Function BuildV8CalcSheet(Optional ByVal targetYear As Long = 0) As Boolean
    Dim w As Worksheet, last As Long
    On Error Resume Next
    Set w = ThisWorkbook.Sheets("GL_V8_CALC")
    On Error GoTo 0
    If w Is Nothing Then Exit Function
    last = w.Cells(w.Rows.Count, 2).End(xlUp).Row
    BuildV8CalcSheet = (last >= 552)
End Function""",
        1,
    ),
]

# ---------------------------------------------------------------------------
# modPJSync.bas
# ---------------------------------------------------------------------------
pjsync_rules = [
    (
        """Private Const ROW_START   As Long = 3""",
        """Private Const ROW_START   As Long = 8""",
        1,
    ),
    (
        """Private Const C_UOM       As Long = 8

Private Const C_AMT       As Long = 9

Private Const C_DEBIT     As Long = 10

Private Const C_MIRROR    As Long = 11

Private Const C_EWT       As Long = 12

Private Const C_VAT       As Long = 13

Private Const C_REF       As Long = 14

Private Const C_COA       As Long = 15

Private Const C_FULLID    As Long = 16""",
        """Private Const C_UOM       As Long = 8

Private Const C_AMT       As Long = 9

Private Const C_DEBIT     As Long = 10

Private Const C_MIRROR    As Long = 9   ' mirror of the amount column

Private Const C_EWT       As Long = 12

Private Const C_VAT       As Long = 13

Private Const C_REF       As Long = 8   ' H = Reference

Private Const C_COA       As Long = 7   ' G = Description / COA

Private Const C_FULLID    As Long = 17  ' Q = Ref Invoice""",
        1,
    ),
    (
        """Private Const CDJ_REF_PNV As Long = 26   ' column Z

Private Const CDJ_NOTE    As Long = 24   ' column X""",
        """Private Const CDJ_REF_PNV As Long = 28   ' column AB

Private Const CDJ_NOTE    As Long = 26   ' column Z""",
        1,
    ),
    # D3: PJ Non-Vat2 is a mirror of PJ; posting from it double-counts.
    (
        """    On Error GoTo EH



    Set sh = PJSheet""",
        """    On Error GoTo EH

    ' GATE2 / decision D3: "PJ Non-Vat2" is a MIRROR of the PJ sheet, not a
    ' separate source of postings. Posting from it would double-count every PJ
    ' transaction. The sync is disabled; lookups and totals stay available.
    Exit Sub

    Set sh = PJSheet""",
        1,
    ),
]

# ---------------------------------------------------------------------------
# Sheet9.cls  (PJ Non-Vat2)
# ---------------------------------------------------------------------------
sheet9_rules = [
    (
        """Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)""",
        """Attribute VB_Customizable = True

Option Explicit

' GATE2: these constants were previously undeclared. Without Option Explicit
' they silently resolved to Empty, every column guard failed, and
' Me.Cells(c.Row, COL_REF) evaluated as Cells(row, 0) and raised an error.
' Values match the verified PJ Non-Vat2 layout (header row 6, data from row 8).
Private Const START_ROW As Long = 8
Private Const COL_TIN As Long = 4
Private Const COL_NAME As Long = 5
Private Const COL_ADDR As Long = 6
Private Const COL_DESC As Long = 7
Private Const COL_COA As Long = 7
Private Const COL_REF As Long = 8
Private Const COL_AMOUNT As Long = 9
Private Const COL_DEBIT As Long = 10
Private Const COL_DEBIT_MIRROR As Long = 9
Private Const COL_DISCOUNT As Long = 11
Private Const COL_EWT As Long = 12
Private Const COL_VAT As Long = 13
Private Const COL_NET As Long = 14
Private Const COL_FULL_ID As Long = 17

Private Sub Worksheet_Change(ByVal Target As Range)""",
        1,
    ),
    # Do not mint PNV refs on a mirror sheet; it would consume the shared sequence.
    (
        """                        If Len(tinVal) > 0 _
                           And Len(nameVal) > 0 _
                           And Len(coaVal) > 0 _
                           And Len(refVal) = 0 Then

                            seqNum = CLng(GetNextSequence())""",
        """                        ' GATE2: reference generation is disabled on this sheet.
                        ' "PJ Non-Vat2" mirrors PJ and must not consume the shared
                        ' PJ reference sequence, nor mint refs that are never posted.
                        If False Then

                            seqNum = CLng(GetNextSequence())""",
        1,
    ),
]


def main():
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    patch("modEngine.bas", engine_rules)
    patch("modPJSync.bas", pjsync_rules)
    patch("Sheet9.cls", sheet9_rules)
    print("\nOK -> %s" % OUT)


if __name__ == "__main__":
    sys.exit(main())
