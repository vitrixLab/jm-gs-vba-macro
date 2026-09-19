Attribute VB_Name = "modGLWorkbookMap"
Option Explicit

' v8.0 exact mapping for Global-Smile_2026-v7.9.
' Posting sources: CDJ, CRJ, GJ. Feeder journals are not posted twice.

Public Const AUDIT_SHEET As String = "GL_AUDIT"
Public Const CALC_SHEET As String = "GL_V8_CALC"
Public Const GL_SHEET As String = "GL"
Public Const TOLERANCE As Double = 0.01

Public Function V8_Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    V8_Norm = UCase$(Application.WorksheetFunction.Trim(Replace(CStr(v), ChrW(160), " ")))
End Function

Public Function V8_Month(ByVal v As Variant) As Long
    Select Case V8_Norm(v)
        Case "JAN", "JANUARY": V8_Month = 1
        Case "FEB", "FEBRUARY": V8_Month = 2
        Case "MAR", "MARCH": V8_Month = 3
        Case "APR", "APRIL": V8_Month = 4
        Case "MAY": V8_Month = 5
        Case "JUN", "JUNE": V8_Month = 6
        Case "JUL", "JULY": V8_Month = 7
        Case "AUG", "AUGUST": V8_Month = 8
        Case "SEP", "SEPTEMBER": V8_Month = 9
        Case "OCT", "OCTOBER": V8_Month = 10
        Case "NOV", "NOVEMBER": V8_Month = 11
        Case "DEC", "DECEMBER": V8_Month = 12
    End Select
End Function

Public Function V8_GLAccounts() As Object
    Dim d As Object, ws As Worksheet, r As Long, a As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    For r = 13 To ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
        a = V8_Norm(ws.Cells(r, 6).Value2)
        If Len(a) > 0 Then If Not d.Exists(a) Then d.Add a, CStr(ws.Cells(r, 6).Value2)
    Next r
    Set V8_GLAccounts = d
End Function

Public Function V8_GJMap(ByVal raw As String) As String
    Dim k As String, d As Object: k = V8_Norm(raw): Set d = V8_GLAccounts()
    Select Case k
        Case "TRANSPORATION AND TRAVEL": V8_GJMap = "Transportation and Travel"
        Case "ADVANCES TO EMPLOYEE": V8_GJMap = "Advances to Employees"
        Case "REPAIRS AND MAINTENANCE": V8_GJMap = "Repair and Maintenance"
        Case "OUTPUT VAT PAYABLE": V8_GJMap = "VAT Payable"
        Case "CLINIC SUPPLIES": V8_GJMap = "Clinic Material and Supplies"
        Case Else: If d.Exists(k) Then V8_GJMap = d(k)
    End Select
End Function

Public Function V8_CDJMap(ByVal col As Long) As String
    Select Case col
        Case 6: V8_CDJMap = "Cash in Bank"
        Case 7: V8_CDJMap = "Input VAT"
        Case 8: V8_CDJMap = "EWT Payable"
        Case 9: V8_CDJMap = "Petty Cash Fund"
        Case 10: V8_CDJMap = "Government Contributions (EE)"
        Case 11: V8_CDJMap = "Government Loans (EE)"
        Case 12: V8_CDJMap = "De Minimis"
        Case 13: V8_CDJMap = "Salaries and Wages"
        Case 14: V8_CDJMap = "Clinic Material and Supplies"
        Case 15: V8_CDJMap = "Rent"
        Case 16: V8_CDJMap = "Gas, Oil, Parking, Toll Fees"
        Case 18: V8_CDJMap = "Professional Fees"
        Case Else: V8_CDJMap = ""
    End Select
End Function

Public Function V8_CRJMap(ByVal col As Long) As String
    Select Case col
        Case 8: V8_CRJMap = "Cash in Bank"
        Case 9: V8_CRJMap = "Excess CWT Over IT"
        Case 10: V8_CRJMap = "VAT Payable"
        Case 11, 12: V8_CRJMap = "Sales"
        Case Else: V8_CRJMap = ""
    End Select
End Function

Public Function V8_Number(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = True
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then ok = False Else V8_Number = CDbl(v)
End Function
