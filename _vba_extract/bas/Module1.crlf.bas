Attribute VB_Name = "Module1"

Option Explicit



' ============================================================

' HELPER: Determine expense type from description

' ============================================================

Function DetermineExpenseType(description As String) As String

    Dim descLower As String

    descLower = LCase(description)

    

    If InStr(descLower, "clinic") > 0 Or InStr(descLower, "dental") > 0 Or InStr(descLower, "medical") > 0 Or InStr(descLower, "material") > 0 Or InStr(descLower, "supplies") > 0 Then

        DetermineExpenseType = "Clinic Materials and Supplies"

    ElseIf InStr(descLower, "rent") > 0 Or InStr(descLower, "lease") > 0 Then

        DetermineExpenseType = "Rental"

    ElseIf InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 Or InStr(descLower, "parking") > 0 Then

        DetermineExpenseType = "Fuel and Oil"

    ElseIf InStr(descLower, "internet") > 0 Or InStr(descLower, "phone") > 0 Or InStr(descLower, "communication") > 0 Or InStr(descLower, "water") > 0 Or InStr(descLower, "light") > 0 Then

        DetermineExpenseType = "Communication"

    ElseIf InStr(descLower, "professional") > 0 Or InStr(descLower, "consulting") > 0 Or InStr(descLower, "lawyer") > 0 Or InStr(descLower, "accounting") > 0 Then

        DetermineExpenseType = "Professional Fees"

    ElseIf InStr(descLower, "repair") > 0 Or InStr(descLower, "maintenance") > 0 Then

        DetermineExpenseType = "Repairs and Maintenance"

    ElseIf InStr(descLower, "supply") > 0 Or InStr(descLower, "office") > 0 Or InStr(descLower, "stationery") > 0 Then

        DetermineExpenseType = "Supplies"

    ElseIf InStr(descLower, "representation") > 0 Or InStr(descLower, "meal") > 0 Or InStr(descLower, "entertainment") > 0 Then

        DetermineExpenseType = "Meals and Entertainment"

    Else

        DetermineExpenseType = "Miscellaneous"

    End If

End Function



' ============================================================

' HELPER: Calculate withholding tax based on expense type

' ============================================================

Function CalculateWithholdingTax(expenseType As String, amount As Double) As Double

    Select Case expenseType

        Case "Clinic Materials and Supplies", "Supplies"

            CalculateWithholdingTax = 0

        Case "Rental"

            CalculateWithholdingTax = amount * 0.05

        Case "Professional Fees"

            CalculateWithholdingTax = amount * 0.1

        Case "Repairs and Maintenance"

            CalculateWithholdingTax = amount * 0.02

        Case "Meals and Entertainment"

            CalculateWithholdingTax = 0

        Case Else

            CalculateWithholdingTax = 0

    End Select

End Function



' ============================================================

' MAIN: Create CDJ entry from PJ or PJ Non-Vat

' ============================================================

Sub CreateCDJFromPurchase(ByVal sourceSheet As Worksheet, ByVal rowNum As Long)

    ' Delegates to the engine writers so CDJ values land where the titles say:
    ' C=Date, E=Particulars, F=Cash In Bank, G=Input VAT, H=WHT, N-S=category,
    ' T=Debit (net + WHT), X=note, Y/Z=reference, AA13/AA14=sequence + row counter.

    If UCase$(sourceSheet.Name) = "PJ NON-VAT" Then

        modEngine.SyncPJNonVatToCDJ sourceSheet, rowNum

    Else

        modEngine.SyncPJToCDJ sourceSheet, rowNum

    End If

End Sub




' ============================================================

' MAIN: Create CRJ entry from SJ

' ============================================================

Sub CreateCRJFromSales(ByVal sourceSheet As Worksheet, ByVal rowNum As Long)

    Dim crjSheet As Worksheet

    Dim customersSheet As Worksheet

    Dim targetRow As Long

    Dim customerName As String

    Dim invoiceNum As String

    Dim amount As Double

    Dim outputVAT As Double

    Dim netSales As Double



    Set crjSheet = ThisWorkbook.Sheets("CRJ")

    

    On Error Resume Next

    Set customersSheet = ThisWorkbook.Sheets("CUSTOMERS")

    On Error GoTo 0

    

    If customersSheet Is Nothing Then

        MsgBox "CUSTOMERS sheet not found. Please create it first.", vbExclamation

        Exit Sub

    End If



    customerName = sourceSheet.Cells(rowNum, "B").value

    invoiceNum = sourceSheet.Cells(rowNum, "E").value

    amount = sourceSheet.Cells(rowNum, "F").value

    outputVAT = sourceSheet.Cells(rowNum, "H").value

    netSales = sourceSheet.Cells(rowNum, "I").value



    If customerName = "" Or amount = 0 Or amount = "" Then Exit Sub



    targetRow = crjSheet.Cells(crjSheet.Rows.count, "A").End(xlUp).row + 1



    With crjSheet

        .Cells(targetRow, "A").value = sourceSheet.Cells(rowNum, "A").value

        .Cells(targetRow, "B").value = customerName

        .Cells(targetRow, "C").value = invoiceNum

        .Cells(targetRow, "D").value = amount + outputVAT

        .Cells(targetRow, "E").Formula = "=D" & targetRow & "*0.107142857"

        .Cells(targetRow, "F").value = 0

        .Cells(targetRow, "G").value = netSales

        

        .Cells(targetRow, "D").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        .Cells(targetRow, "E").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        .Cells(targetRow, "F").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        .Cells(targetRow, "G").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        .Cells(targetRow, "H").value = "Auto-generated from SJ Row " & rowNum

    End With

End Sub



' ============================================================

' BACKFILL: Process all existing rows

' ============================================================

Sub BackfillAllAutomation()

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






