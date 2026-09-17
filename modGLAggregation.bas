Attribute VB_Name = "modGLAggregation"
Option Explicit

' v7.9.3 GL hardening engine.
' Accounting flow:
'   posted journals -> normalized account/month -> debit/credit totals -> running GL balance
'
' This module deliberately does not use Graphify or numeric-column guessing for accounting.

Private Const GL_SHEET As String = "GL"
Private Const AUDIT_SHEET As String = "GL_AUDIT"
Private Const COA_SHEET As String = "SUPPLIERS DATA"
Private Const TOLERANCE As Double = 0.01

Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Application.WorksheetFunction.Trim(s)
    Norm = UCase$(s)
End Function

Private Function IsSupportedSource(ByVal sheetName As String) As Boolean
    Select Case UCase$(sheetName)
        Case "PJ", "PJ NON-VAT", "CDJ", "SJ", "CRJ", "GJ"
            IsSupportedSource = True
        Case Else
            IsSupportedSource = False
    End Select
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByVal labels As Variant, Optional ByVal maxRows As Long = 40, Optional ByVal maxCols As Long = 40) As Long
    Dim r As Long, c As Long, i As Long, target As String, cellText As String
    For r = 1 To Application.Min(maxRows, ws.UsedRange.Rows.Count)
        For c = 1 To Application.Min(maxCols, ws.UsedRange.Columns.Count)
            cellText = Norm(ws.Cells(r, c).Value2)
            If Len(cellText) > 0 Then
                For i = LBound(labels) To UBound(labels)
                    target = Norm(labels(i))
                    If cellText = target Then
                        FindHeader = c
                        Exit Function
                    End If
                Next i
            End If
        Next c
    Next r
    FindHeader = 0
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal requiredLabels As Variant, Optional ByVal maxRows As Long = 40, Optional ByVal maxCols As Long = 40) As Long
    Dim r As Long, c As Long, i As Long, hits As Long, txt As String
    For r = 1 To Application.Min(maxRows, ws.UsedRange.Rows.Count)
        hits = 0
        For c = 1 To Application.Min(maxCols, ws.UsedRange.Columns.Count)
            txt = Norm(ws.Cells(r, c).Value2)
            For i = LBound(requiredLabels) To UBound(requiredLabels)
                If txt = Norm(requiredLabels(i)) Then hits = hits + 1: Exit For
            Next i
        Next c
        If hits = UBound(requiredLabels) - LBound(requiredLabels) + 1 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
    FindHeaderRow = 0
End Function

Private Function ToNumber(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        ToNumber = 0
    ElseIf IsNumeric(v) Then
        ToNumber = CDbl(v)
    Else
        ToNumber = 0
    End If
End Function

Private Function AccountDictionary() As Object
    Dim d As Object, ws As Worksheet, lastRow As Long, r As Long, acct As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(COA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For r = 5 To lastRow
        acct = Norm(ws.Cells(r, 6).Value2)
        If Len(acct) > 0 Then
            If Not d.Exists(acct) Then d.Add acct, CStr(ws.Cells(r, 6).Value2)
        End If
    Next r
    Set AccountDictionary = d
    Exit Function
Fail:
    Set AccountDictionary = d
End Function

Public Function GetAccountNameFromTIN(ByVal tin As String) As String
    ' Fixed v7.9.2 defect: no undefined key variable.
    Dim d As Object, ws As Worksheet, lastRow As Long, r As Long
    Dim needle As String
    needle = Norm(tin)
    If Len(needle) = 0 Then Exit Function
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(COA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For r = 5 To lastRow
        If Norm(ws.Cells(r, 3).Value2) = needle Then
            GetAccountNameFromTIN = CStr(ws.Cells(r, 6).Value2)
            Exit Function
        End If
    Next r
Fail:
End Function

Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    Dim d As Object, key As String
    Set d = AccountDictionary()
    key = Norm(coaCode)
    If d.Exists(key) Then GetAccountFromCOA = CStr(d(key))
End Function

Private Function JournalAccount(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal accountCol As Long, ByVal tinCol As Long, ByVal nameCol As Long) As String
    Dim s As String
    If accountCol > 0 Then s = Trim$(CStr(ws.Cells(rowNum, accountCol).Value2))
    If Len(s) = 0 And tinCol > 0 Then s = GetAccountNameFromTIN(CStr(ws.Cells(rowNum, tinCol).Value2))
    If Len(s) = 0 And nameCol > 0 Then s = Trim$(CStr(ws.Cells(rowNum, nameCol).Value2))
    JournalAccount = s
End Function

Private Function ParseDateValue(ByVal v As Variant) As Date
    If IsDate(v) Then ParseDateValue = CDate(v) Else ParseDateValue = 0
End Function

Private Function IsPostedRow(ByVal ws As Worksheet, ByVal r As Long, ByVal statusCol As Long) As Boolean
    Dim s As String
    If statusCol = 0 Then
        ' Existing journal sheets do not expose a universal Posted header; in that case
        ' a populated journal row is the supported source population.
        IsPostedRow = True
    Else
        s = Norm(ws.Cells(r, statusCol).Value2)
        IsPostedRow = (s = "POSTED" Or s = "YES" Or s = "TRUE" Or s = "1")
    End If
End Function

Public Function CalculateMonthlyNet(ByVal accountName As String, ByVal monthNumber As Long, ByVal yearNumber As Long) As Double
    Dim totals As Variant
    totals = BuildGLMonthlyTotals(accountName, monthNumber, yearNumber)
    CalculateMonthlyNet = CDbl(totals(0)) - CDbl(totals(1))
End Function

Public Function BuildGLMonthlyTotals(ByVal accountName As String, ByVal monthNumber As Long, ByVal yearNumber As Long) As Variant
    Dim sources As Variant, source As Variant, ws As Worksheet
    Dim headerRow As Long, dateCol As Long, debitCol As Long, creditCol As Long
    Dim acctCol As Long, tinCol As Long, nameCol As Long, statusCol As Long
    Dim r As Long, lastRow As Long, dt As Date, acct As String
    Dim debitTotal As Double, creditTotal As Double
    Dim result(0 To 1) As Double

    sources = Array("PJ", "PJ Non-Vat", "CDJ", "SJ", "CRJ", "GJ")
    For Each source In sources
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(source))
        On Error GoTo 0
        If Not ws Is Nothing Then
            headerRow = FindHeaderRow(ws, Array("Debit", "Credit"), 40, 60)
            If headerRow = 0 Then GoTo NextSource
            dateCol = FindHeader(ws, Array("Date", "Transaction Date", "Journal Date"), 40, 60)
            debitCol = FindHeader(ws, Array("Debit"), 40, 60)
            creditCol = FindHeader(ws, Array("Credit"), 40, 60)
            acctCol = FindHeader(ws, Array("Account", "Account Title", "Chart of Account", "CHART OF ACCOUNT"), 40, 60)
            tinCol = FindHeader(ws, Array("Suppliers' TIN", "Customers' TIN", "TIN#", "TIN"), 40, 60)
            nameCol = FindHeader(ws, Array("Vendors' Name", "Customers' Name", "Particulars", "Explanation"), 40, 60)
            statusCol = FindHeader(ws, Array("Status", "Posted"), 40, 60)
            lastRow = ws.Cells(ws.Rows.Count, IIf(dateCol > 0, dateCol, debitCol)).End(xlUp).Row
            For r = headerRow + 1 To lastRow
                If IsPostedRow(ws, r, statusCol) Then
                    If dateCol > 0 Then dt = ParseDateValue(ws.Cells(r, dateCol).Value2) Else dt = DateSerial(yearNumber, monthNumber, 1)
                    If (dateCol = 0 Or (Year(dt) = yearNumber And Month(dt) = monthNumber)) Then
                        acct = Norm(JournalAccount(ws, r, acctCol, tinCol, nameCol))
                        If acct = Norm(accountName) Then
                            debitTotal = debitTotal + ToNumber(ws.Cells(r, debitCol).Value2)
                            creditTotal = creditTotal + ToNumber(ws.Cells(r, creditCol).Value2)
                        End If
                    End If
                End If
            Next r
        End If
NextSource:
    Next source
    result(0) = debitTotal
    result(1) = creditTotal
    BuildGLMonthlyTotals = result
End Function

Public Function BuildGLOpeningBalance(ByVal accountName As String, Optional ByVal openingBalance As Double = 0) As Double
    ' Opening balance is an explicit input. No silent inference from current GL cells.
    BuildGLOpeningBalance = openingBalance
End Function

Public Function CalculateEndingBalanceFull(ByVal accountName As String, ByVal yearNumber As Long, Optional ByVal openingBalance As Double = 0) As Double
    Dim m As Long, net As Double, balance As Double
    balance = BuildGLOpeningBalance(accountName, openingBalance)
    For m = 1 To 12
        net = CalculateMonthlyNet(accountName, m, yearNumber)
        balance = balance + net
    Next m
    CalculateEndingBalanceFull = balance
End Function

Public Function ValidateGLConsistency(Optional ByVal yearNumber As Long = 0) As Boolean
    Dim sources As Variant, source As Variant, ws As Worksheet
    Dim headerRow As Long, debitCol As Long, creditCol As Long, statusCol As Long, dateCol As Long
    Dim r As Long, lastRow As Long, dt As Date, totalD As Double, totalC As Double
    Dim mappedD As Double, mappedC As Double, unmapped As Double
    Dim d As Object, acct As String, report As String

    If yearNumber = 0 Then yearNumber = Year(Date)
    Set d = AccountDictionary()
    sources = Array("PJ", "PJ Non-Vat", "CDJ", "SJ", "CRJ", "GJ")
    For Each source In sources
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(source))
        On Error GoTo 0
        If Not ws Is Nothing Then
            headerRow = FindHeaderRow(ws, Array("Debit", "Credit"), 40, 60)
            If headerRow = 0 Then GoTo NextValidationSource
            debitCol = FindHeader(ws, Array("Debit"), 40, 60)
            creditCol = FindHeader(ws, Array("Credit"), 40, 60)
            statusCol = FindHeader(ws, Array("Status", "Posted"), 40, 60)
            dateCol = FindHeader(ws, Array("Date", "Transaction Date", "Journal Date"), 40, 60)
            lastRow = ws.Cells(ws.Rows.Count, IIf(dateCol > 0, dateCol, debitCol)).End(xlUp).Row
            For r = headerRow + 1 To lastRow
                If IsPostedRow(ws, r, statusCol) Then
                    If dateCol > 0 Then
                        dt = ParseDateValue(ws.Cells(r, dateCol).Value2)
                        If Year(dt) <> yearNumber Then GoTo NextValidationRow
                    End If
                    totalD = totalD + ToNumber(ws.Cells(r, debitCol).Value2)
                    totalC = totalC + ToNumber(ws.Cells(r, creditCol).Value2)
                    acct = Norm(JournalAccount(ws, r, FindHeader(ws, Array("Account", "Account Title", "Chart of Account", "CHART OF ACCOUNT"), 40, 60), FindHeader(ws, Array("Suppliers' TIN", "Customers' TIN", "TIN#", "TIN"), 40, 60), FindHeader(ws, Array("Vendors' Name", "Customers' Name", "Particulars", "Explanation"), 40, 60)))
                    If Len(acct) = 0 Or Not d.Exists(acct) Then
                        unmapped = unmapped + ToNumber(ws.Cells(r, debitCol).Value2) + ToNumber(ws.Cells(r, creditCol).Value2)
                    Else
                        mappedD = mappedD + ToNumber(ws.Cells(r, debitCol).Value2)
                        mappedC = mappedC + ToNumber(ws.Cells(r, creditCol).Value2)
                    End If
                End If
NextValidationRow:
            Next r
        End If
NextValidationSource:
    Next source

    report = "Year=" & CStr(yearNumber) & vbCrLf & _
             "Total Debit=" & Format$(totalD, "0.00") & vbCrLf & _
             "Total Credit=" & Format$(totalC, "0.00") & vbCrLf & _
             "Difference=" & Format$(totalD - totalC, "0.00") & vbCrLf & _
             "Mapped Debit=" & Format$(mappedD, "0.00") & vbCrLf & _
             "Mapped Credit=" & Format$(mappedC, "0.00") & vbCrLf & _
             "Unmapped Activity=" & Format$(unmapped, "0.00")
    WriteAudit report, (Abs(totalD - totalC) <= TOLERANCE And unmapped <= TOLERANCE)
    ValidateGLConsistency = (Abs(totalD - totalC) <= TOLERANCE And unmapped <= TOLERANCE)
End Function

Public Sub RefreshGL(Optional ByVal yearNumber As Long = 0)
    If yearNumber = 0 Then yearNumber = Year(Date)
    If Not ValidateGLConsistency(yearNumber) Then
        MsgBox "GL refresh stopped: validation did not pass. Review GL_AUDIT.", vbExclamation
        Exit Sub
    End If
    ' v7.9.3 intentionally stops before overwriting GL until the mapping gate is proven.
    MsgBox "GL validation passed. Workbook GL write phase is ready for the next controlled gate.", vbInformation
End Sub

Public Sub RefreshAllGL(Optional ByVal yearNumber As Long = 0)
    RefreshGL yearNumber
End Sub

Private Sub WriteAudit(ByVal report As String, ByVal passed As Boolean)
    Dim ws As Worksheet, lines As Variant, i As Long, nextRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = AUDIT_SHEET
    End If
    On Error GoTo 0
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = IIf(passed, "PASS", "HOLD")
    lines = Split(report, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        ws.Cells(nextRow + i, 3).Value = lines(i)
    Next i
End Sub
