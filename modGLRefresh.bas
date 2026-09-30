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
Private Const GL_ACCOUNT_BLOCKS As Long = 46
Private Const GL_MONTHS As Long = 12

Public Sub RefreshGLIntoSheet(Optional ByVal yearNumber As Long = 2026)
    Dim d As Object, gd As Double, gc As Double, um As Double, inv As Long, detail As String
    If Not ValidateV8Structure() Then
        AppendAudit "HOLD", "v8.2 | structure validation failed", "GL not overwritten."
        MsgBox "v8.2 HOLD: workbook structure does not match the reviewed mapping.", vbExclamation
        Exit Sub
    End If
    Set d = Matrix(yearNumber, gd, gc, um, inv, detail)
    If inv > 0 Or um > TOLERANCE Or Abs(gd - gc) > TOLERANCE Then
        AppendAudit "HOLD", "v8.2 | CDJ+CRJ+GJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
            " | Difference=" & Fmt2(gd - gc) & " | Unmapped=" & Fmt2(um) & " | Invalid=" & inv, detail
        MsgBox "v8.2 HOLD: reconciliation/mapping gate failed. GL not overwritten.", vbExclamation
        Exit Sub
    End If
    WriteGLSheet d
    RefreshCalcFromMatrix d
    AppendAudit "PASS", "v8.2 | CDJ+CRJ+GJ | Debit=" & Fmt2(gd) & " | Credit=" & Fmt2(gc) & _
        " | Difference=" & Fmt2(gd - gc) & " | Unmapped=" & Fmt2(um) & " | Invalid=" & inv & _
        " | 46x12 GL filled | ending-balance matrix refreshed", ""
    MsgBox "v8.2 PASS: GL refreshed for all 46 accounts x 12 months from CDJ+CRJ+GJ.", vbInformation
End Sub

Public Sub RefreshGL(Optional ByVal yearNumber As Long = 2026)
    RefreshGLIntoSheet yearNumber
End Sub

Public Sub RefreshAllGL(Optional ByVal yearNumber As Long = 2026)
    RefreshGLIntoSheet yearNumber
End Sub

' --- write every GL block row in place: G = debit, H = credit (0 when none)
Private Sub WriteGLSheet(ByVal d As Object)
    Dim ws As Worksheet, b As Long, m As Long, r As Long, k As String, x As Variant
    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    If Not ValidateGLWriteGrid(ws) Then Err.Raise vbObjectError + 832, "RefreshGL", "GL write grid is not the expected 46 x 12 layout."
    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    For b = 0 To GL_ACCOUNT_BLOCKS - 1
        For m = 1 To GL_MONTHS
            r = GL_FIRST_ROW + b * GL_BLOCK_STRIDE + (m - 1)
            k = V8_Norm(CStr(ws.Cells(r, 6).Value2)) & "|" & m
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            ws.Cells(r, GL_DEBIT_COL).Value2 = CDbl(x(0))
            ws.Cells(r, GL_CREDIT_COL).Value2 = CDbl(x(1))
        Next m
    Next b
    Application.Calculate
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ValidateGLWriteGrid(ByVal ws As Worksheet) As Boolean
    Dim b As Long, m As Long, r As Long
    For b = 0 To GL_ACCOUNT_BLOCKS - 1
        r = GL_FIRST_ROW + b * GL_BLOCK_STRIDE
        If Len(Trim$(CStr(ws.Cells(r, 6).Value2))) = 0 Then Exit Function
        For m = 1 To GL_MONTHS
            r = GL_FIRST_ROW + b * GL_BLOCK_STRIDE + (m - 1)
            If ws.Cells(r, GL_DEBIT_COL).MergeCells Or ws.Cells(r, GL_CREDIT_COL).MergeCells Then Exit Function
        Next m
    Next b
    ValidateGLWriteGrid = True
End Function

' --- refresh the kept GL_V8_CALC skeleton from the same matrix
Private Sub RefreshCalcFromMatrix(ByVal d As Object)
    Dim ws As Worksheet, b As Long, m As Long, r As Long, k As String, x As Variant
    Dim title As String, endingBalance As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.Range("A4:E555").MergeCells Then Err.Raise vbObjectError + 833, "RefreshGL", "GL_V8_CALC contains merged cells in the 552-row output range."
    ws.Range("A4:E555").ClearContents
    r = 4
    For b = 0 To 45
        title = CStr(ThisWorkbook.Worksheets(GL_SHEET).Cells(GL_FIRST_ROW + b * GL_BLOCK_STRIDE, 6).Value2)
        endingBalance = 0#
        For m = 1 To 12
            k = V8_Norm(title) & "|" & m
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            endingBalance = endingBalance + CDbl(x(0)) - CDbl(x(1))
            ws.Cells(r, 1).Value2 = title
            ws.Cells(r, 2).Value2 = m
            ws.Cells(r, 3).Value2 = CDbl(x(0))
            ws.Cells(r, 4).Value2 = CDbl(x(1))
            ws.Cells(r, 5).Value2 = endingBalance
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


