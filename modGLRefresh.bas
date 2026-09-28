Attribute VB_Name = "modGLRefresh"
Option Explicit

' v8.1 RefreshGL — fills the EXISTING GL sheet (no new sheets) from CDJ + CRJ.
' Scope: CDJ + CRJ only (GJ excluded). All 46 blocks x 12 months written; 0 if none.
' GL layout: blocks of 12 rows from row 13, stride 13 (blank separator).
' Col B = month, col F = title, col G = Debit, col H = Credit.
' CDJ F(6):S(19) signed; T(20) is check col, NOT a posting. S(19) text label
' amount = row cash-out (T, else -F) minus signed F:R allocation.
' CRJ: H(8) = Cash in Bank debit; I(9):L(12) credits; month from Entry-Log date.

Private Const GL_FIRST_ROW As Long = 13
Private Const GL_BLOCK_STRIDE As Long = 13
Private Const GL_DEBIT_COL As Long = 7
Private Const GL_CREDIT_COL As Long = 8

Public Sub RefreshGLIntoSheet(Optional ByVal yearNumber As Long = 2026)
    Dim d As Object, gd As Double, gc As Double, um As Double, inv As Long, detail As String
    If Not ValidateV8Structure() Then
        MsgBox "v8.1 HOLD: workbook structure does not match the reviewed mapping.", vbExclamation
        Exit Sub
    End If
    Set d = AggregateCDJCRJ(yearNumber, gd, gc, um, inv, detail)
    If inv > 0 Then
        AppendAudit "HOLD", "v8.1 | CDJ+CRJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
            " | Difference=" & Fmt2(gd - gc) & " | Invalid=" & inv, detail
        MsgBox "v8.1 HOLD: invalid numeric/date activity. Review GL_AUDIT.", vbExclamation
        Exit Sub
    End If
    If Abs(gd - gc) > TOLERANCE Then
        AppendAudit "HOLD", "v8.1 | CDJ+CRJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
            " | Difference=" & Fmt2(gd - gc) & " | Unmapped=" & Fmt2(um), detail
        MsgBox "v8.1 HOLD: CDJ+CRJ out of balance (" & Fmt2(gd - gc) & "). GL not overwritten.", vbExclamation
        Exit Sub
    End If
    If um > TOLERANCE Then
        AppendAudit "HOLD", "v8.1 | CDJ+CRJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
            " | Unmapped=" & Fmt2(um), detail
        MsgBox "v8.1 HOLD: unmapped posting activity. Review GL_AUDIT.", vbExclamation
        Exit Sub
    End If
    WriteGLSheet d
    RefreshCalcFromMatrix d
    AppendAudit "PASS", "v8.1 | CDJ+CRJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
        " | Difference=" & Fmt2(gd - gc) & " | Unmapped=" & Fmt2(um) & " | Invalid=" & inv & _
        " | 46x12 GL filled (GJ excluded)", ""
    MsgBox "v8.1 PASS: GL filled for all 46 accounts x 12 months from CDJ+CRJ.", vbInformation
End Sub

Public Sub RefreshGL(Optional ByVal yearNumber As Long = 2026)
    RefreshGLIntoSheet yearNumber
End Sub


Public Sub RefreshAllGL(Optional ByVal yearNumber As Long = 2026)
    RefreshGLIntoSheet yearNumber
End Sub

' --- aggregate CDJ + CRJ into dict "NORMED_ACCT|m" -> Array(debit, credit)
Private Function AggregateCDJCRJ(ByVal yr As Long, ByRef gd As Double, ByRef gc As Double, _
    ByRef um As Double, ByRef inv As Long, ByRef detail As String) As Object
    Dim d As Object, ws As Worksheet, r As Long, c As Long, m As Long
    Dim a As String, v As Double, ok As Boolean, signed As Double, cashOut As Double
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets("CDJ")
    m = 0
    For r = 15 To ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
        If V8_Month(ws.Cells(r, 3).Value2) > 0 Then m = V8_Month(ws.Cells(r, 3).Value2)
        If V8_Norm(ws.Cells(r, 5).Value2) = "TOTAL" Then GoTo NextCDJ
        If m >= 1 And m <= 12 And Len(Trim$(CStr(ws.Cells(r, 5).Value2))) > 0 Then
            signed = 0#
            For c = 6 To 18
                ok = True: v = V8_Number(ws.Cells(r, c).Value2, ok)
                If Not ok Then inv = inv + 1
                If ok And Abs(v) > TOLERANCE Then
                    a = V8_CDJMap(c)
                    If Len(a) = 0 Then
                        um = um + Abs(v)
                        detail = detail & "CDJ!" & ws.Cells(r, c).Address(False, False) & " ambiguous" & vbCrLf
                    ElseIf v >= 0 Then
                        PutM d, a, m, v, 0#: gd = gd + v: signed = signed + v
                    Else
                        PutM d, a, m, 0#, -v: gc = gc - v: signed = signed + v
                    End If
                End If
            Next c
            a = Trim$(CStr(ws.Cells(r, 19).Value2))
            If Len(a) > 0 And UCase$(a) <> "TOTAL" And UCase$(a) <> "SUNDRY ACCOUNT" Then
                ok = True: cashOut = V8_Number(ws.Cells(r, 20).Value2, ok)
                If Abs(cashOut) <= TOLERANCE Then
                    ok = True: v = V8_Number(ws.Cells(r, 6).Value2, ok)
                    If ok Then cashOut = -v
                End If
                v = cashOut - signed
                If Abs(v) > TOLERANCE Then
                    If V8_GJMap(a) <> "" Then a = V8_GJMap(a) _
                    ElseIf GLTitleExists(a) Then a = GLTitleOf(a) _
                    Else um = um + Abs(v): detail = detail & "CDJ!" & _
                        ws.Cells(r, 19).Address(False, False) & " -> " & a & vbCrLf: GoTo NextCDJ
                    If v >= 0 Then PutM d, a, m, v, 0#: gd = gd + v _
                    Else PutM d, a, m, 0#, -v: gc = gc - v
                End If
            End If
        End If
NextCDJ:
    Next r
    Set ws = ThisWorkbook.Worksheets("CRJ")
    Dim dt As Date, mm As Long
    For r = 10 To ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
        If V8_Norm(ws.Cells(r, 6).Value2) = "TOTAL" Then GoTo NextCRJ
        If Len(Trim$(CStr(ws.Cells(r, 7).Value2))) = 0 Then GoTo NextCRJ
        dt = EntryLogDate(CStr(ws.Cells(r, 15).Value2))
        If dt = 0 Then inv = inv + 1: GoTo NextCRJ
        If Year(dt) <> yr Then GoTo NextCRJ
        mm = Month(dt)
        ok = True: v = V8_Number(ws.Cells(r, 8).Value2, ok)
        If Not ok Then inv = inv + 1: GoTo NextCRJ
        If Abs(v) > TOLERANCE Then PutM d, "Cash in Bank", mm, v, 0#: gd = gd + v
        For c = 9 To 12
            ok = True: v = V8_Number(ws.Cells(r, c).Value2, ok)
            If Not ok Then inv = inv + 1
            If ok And Abs(v) > TOLERANCE Then
                a = V8_CRJMap(c)
                If Len(a) = 0 Then
                    um = um + Abs(v)
                    detail = detail & "CRJ!" & ws.Cells(r, c).Address(False, False) & " unmapped" & vbCrLf
                Else
                    PutM d, a, mm, 0#, v: gc = gc + v
                End If
            End If
        Next c
NextCRJ:
    Next r
    Set AggregateCDJCRJ = d
End Function

Private Sub PutM(ByVal d As Object, ByVal acct As String, ByVal m As Long, ByVal db As Double, ByVal cr As Double)
    Dim k As String, x As Variant
    k = V8_Norm(acct) & "|" & m
    If Not d.Exists(k) Then d.Add k, Array(0#, 0#)
    x = d(k): x(0) = x(0) + db: x(1) = x(1) + cr: d(k) = x
End Sub

' --- write every GL block row in place: G = debit, H = credit (0 when none)
Private Sub WriteGLSheet(ByVal d As Object)
    Dim ws As Worksheet, b As Long, m As Long, r As Long, k As String, x As Variant
    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    Application.ScreenUpdating = False
    For b = 0 To 45
        For m = 1 To 12
            r = GL_FIRST_ROW + b * GL_BLOCK_STRIDE + (m - 1)
            k = V8_Norm(CStr(ws.Cells(r, 6).Value2)) & "|" & m
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            ws.Cells(r, GL_DEBIT_COL).Value2 = CDbl(x(0))
            ws.Cells(r, GL_CREDIT_COL).Value2 = CDbl(x(1))
        Next m
    Next b
    Application.ScreenUpdating = True
End Sub

' --- refresh the kept GL_V8_CALC skeleton from the same matrix
Private Sub RefreshCalcFromMatrix(ByVal d As Object)
    Dim ws As Worksheet, b As Long, m As Long, r As Long, k As String, x As Variant
    Dim title As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    r = 4
    For b = 0 To 45
        title = CStr(ThisWorkbook.Worksheets(GL_SHEET).Cells(GL_FIRST_ROW + b * GL_BLOCK_STRIDE, 6).Value2)
        For m = 1 To 12
            k = V8_Norm(title) & "|" & m
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            ws.Cells(r, 1).Value2 = title
            ws.Cells(r, 2).Value2 = m
            ws.Cells(r, 3).Value2 = CDbl(x(0))
            ws.Cells(r, 4).Value2 = CDbl(x(1))
            ws.Cells(r, 5).Value2 = CDbl(x(0)) - CDbl(x(1))
            r = r + 1
        Next m
    Next b
End Sub

Private Sub AppendAudit(ByVal status As String, ByVal summary As String, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): ws.Name = AUDIT_SHEET
    On Error GoTo 0
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = status
    ws.Cells(r, 3).Value = summary
    ws.Cells(r, 4).Value = detail
End Sub

Private Function EntryLogDate(ByVal s As String) As Date
    Dim p As Long
    p = InStrRev(s, " | "): If p > 0 Then s = Mid$(s, p + 3)
    If IsDate(s) Then EntryLogDate = CDate(s)
End Function

Private Function Fmt2(ByVal v As Double) As String
    Fmt2 = Format$(v, "0.00")
End Function

Private Function GLTitleExists(ByVal raw As String) As Boolean
    GLTitleExists = V8_GLAccounts().Exists(V8_Norm(raw))
End Function

Private Function GLTitleOf(ByVal raw As String) As String
    GLTitleOf = V8_GLAccounts()(V8_Norm(raw))
End Function


