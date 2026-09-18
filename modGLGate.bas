Attribute VB_Name = "modGLGate"
Option Explicit

' Fail-closed gate for the v7.9.3 GL engine.
' This gate is intentionally separate from Graphify and from journal calculations.

Private Const AUDIT_SHEET As String = "GL_AUDIT"
Private Const MAX_HEADER_ROWS As Long = 40
Private Const MAX_HEADER_COLS As Long = 60

Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Application.WorksheetFunction.Trim(s)
    Norm = UCase$(s)
End Function

Private Function HeaderCol(ByVal ws As Worksheet, ByVal labels As Variant) As Long
    Dim r As Long, c As Long, i As Long, txt As String
    For r = 1 To Application.Min(MAX_HEADER_ROWS, ws.UsedRange.Rows.Count)
        For c = 1 To Application.Min(MAX_HEADER_COLS, ws.UsedRange.Columns.Count)
            txt = Norm(ws.Cells(r, c).Value2)
            For i = LBound(labels) To UBound(labels)
                If txt = Norm(labels(i)) Then HeaderCol = c: Exit Function
            Next i
        Next c
    Next r
End Function

Private Function HeaderRow(ByVal ws As Worksheet, ByVal labels As Variant) As Long
    Dim r As Long, c As Long, i As Long, hits As Long, txt As String
    For r = 1 To Application.Min(MAX_HEADER_ROWS, ws.UsedRange.Rows.Count)
        hits = 0
        For c = 1 To Application.Min(MAX_HEADER_COLS, ws.UsedRange.Columns.Count)
            txt = Norm(ws.Cells(r, c).Value2)
            For i = LBound(labels) To UBound(labels)
                If txt = Norm(labels(i)) Then hits = hits + 1: Exit For
            Next i
        Next c
        If hits = UBound(labels) - LBound(labels) + 1 Then HeaderRow = r: Exit Function
    Next r
End Function

Public Function ValidateWorkbookMapping(Optional ByVal writeAudit As Boolean = True) As Boolean
    Dim sources As Variant, source As Variant, ws As Worksheet
    Dim hr As Long, dateCol As Long, debitCol As Long, creditCol As Long, acctCol As Long
    Dim missing As String, ok As Boolean
    ok = True
    sources = Array("PJ", "PJ Non-Vat", "CDJ", "SJ", "CRJ", "GJ")

    For Each source In sources
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(source))
        On Error GoTo 0
        If ws Is Nothing Then
            ok = False
            missing = missing & CStr(source) & ": sheet missing; "
        Else
            hr = HeaderRow(ws, Array("Debit", "Credit"))
            debitCol = HeaderCol(ws, Array("Debit"))
            creditCol = HeaderCol(ws, Array("Credit"))
            dateCol = HeaderCol(ws, Array("Date", "Transaction Date", "Journal Date"))
            acctCol = HeaderCol(ws, Array("Account", "Account Title", "Chart of Account", "CHART OF ACCOUNT"))
            If hr = 0 Or debitCol = 0 Or creditCol = 0 Then
                ok = False
                missing = missing & CStr(source) & ": Debit/Credit mapping missing; "
            End If
            ' Date or an explicitly documented source-specific period rule is required.
            If dateCol = 0 Then
                ok = False
                missing = missing & CStr(source) & ": Date mapping missing; "
            End If
            ' Account can be resolved through a source-specific TIN/name mapping, but that
            ' mapping must be proven before the GL write gate is opened.
            If acctCol = 0 Then
                If HeaderCol(ws, Array("Suppliers' TIN", "Customers' TIN", "TIN#", "TIN")) = 0 And _
                   HeaderCol(ws, Array("Vendors' Name", "Customers' Name", "Particulars", "Explanation")) = 0 Then
                    ok = False
                    missing = missing & CStr(source) & ": Account/TIN/name mapping missing; "
                End If
            End If
        End If
    Next source

    If writeAudit Then WriteGateAudit ok, missing
    ValidateWorkbookMapping = ok
End Function

Private Sub WriteGateAudit(ByVal passed As Boolean, ByVal details As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = AUDIT_SHEET
    End If
    On Error GoTo 0
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = IIf(passed, "MAPPING PASS", "MAPPING HOLD")
    ws.Cells(r, 3).Value = details
End Sub

Public Sub RunV793Gate()
    If ValidateWorkbookMapping(True) Then
        MsgBox "v7.9.3 mapping gate PASSED. Proceed to extraction/reconciliation tests.", vbInformation
    Else
        MsgBox "v7.9.3 mapping gate HOLD. Review GL_AUDIT before implementing GL writes.", vbExclamation
    End If
End Sub
