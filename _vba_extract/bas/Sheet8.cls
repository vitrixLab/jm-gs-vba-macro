Attribute VB_Name = "Sheet8"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Const START_ROW As Long = 10



' ============================================================

' FIND REFERENCE IN CRJ (Fast - uses Find)

' ============================================================

Private Function FindReferenceInCRJ(invoiceNum As String) As Long

    Dim wsCRJ As Worksheet, found As Range

    On Error Resume Next

    Set wsCRJ = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0

    If wsCRJ Is Nothing Then FindReferenceInCRJ = 0: Exit Function

    

    Set found = wsCRJ.Columns(3).Find(What:=invoiceNum, LookIn:=xlValues, LookAt:=xlWhole)

    ' Column 3 = C (Invoice #)

    

    If Not found Is Nothing Then

        FindReferenceInCRJ = found.row

    Else

        FindReferenceInCRJ = 0

    End If

End Function



' ============================================================

' SYNC SJ TO CRJ

' ============================================================

Private Sub SyncToCRJ(sourceRow As Long)

    Dim wsSJ As Worksheet, wsCRJ As Worksheet

    Dim invoiceNum As Variant, targetRow As Long

    Dim customerName As String

    Dim amount As Double, outputVAT As Double, netSales As Double

    

    On Error GoTo ErrorHandler

    

    Set wsSJ = Me

    Set wsCRJ = ThisWorkbook.Sheets("CRJ")

    If wsCRJ Is Nothing Then Exit Sub

    

    invoiceNum = wsSJ.Cells(sourceRow, 9).value  ' I = Invoice #

    If invoiceNum = "" Then Exit Sub

    

    customerName = wsSJ.Cells(sourceRow, 6).value   ' F = Customer Name

    amount = wsSJ.Cells(sourceRow, 10).value        ' J = Amount

    outputVAT = wsSJ.Cells(sourceRow, 12).value     ' L = Output VAT

    netSales = wsSJ.Cells(sourceRow, 14).value      ' N = Net Sales

    

    If customerName = "" Or amount = 0 Or amount = "" Then Exit Sub

    

    targetRow = FindReferenceInCRJ(invoiceNum)

    

    If targetRow > 0 Then

        ' UPDATE existing row

        With wsCRJ

            .Cells(targetRow, 1).value = wsSJ.Cells(sourceRow, 4).value    ' A = Date

            .Cells(targetRow, 2).value = customerName                      ' B = Customer

            .Cells(targetRow, 3).value = invoiceNum                        ' C = Invoice #

            .Cells(targetRow, 4).value = amount + outputVAT                ' D = Cash in Bank

            .Cells(targetRow, 5).value = outputVAT                         ' E = Output VAT

            .Cells(targetRow, 6).value = 0                                 ' F = Exempt Sales

            .Cells(targetRow, 7).value = netSales                          ' G = Sales

            .Cells(targetRow, 8).value = "Updated: " & Now & " | Invoice: " & invoiceNum

        End With

    Else

        ' INSERT new row

        targetRow = wsCRJ.Cells(wsCRJ.Rows.count, 1).End(xlUp).row + 1

        If targetRow < 10 Then targetRow = 10

        

        With wsCRJ

            .Cells(targetRow, 1).value = wsSJ.Cells(sourceRow, 4).value

            .Cells(targetRow, 2).value = customerName

            .Cells(targetRow, 3).value = invoiceNum

            .Cells(targetRow, 4).value = amount + outputVAT

            .Cells(targetRow, 5).value = outputVAT

            .Cells(targetRow, 6).value = 0

            .Cells(targetRow, 7).value = netSales

            .Cells(targetRow, 8).value = "Auto-generated from SJ Row " & sourceRow & " | Invoice: " & invoiceNum

            

            .Cells(targetRow, 4).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, 5).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, 6).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        End With

    End If

    

    Exit Sub

ErrorHandler:

End Sub



' ============================================================

' MAIN TRIGGER

' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range, cell As Range

    Dim seqNum As Long, i As Long, lastRow As Long, maxSeq As Long



    Set rng = Intersect(Target, Union(Me.Range("F:F"), Me.Range("G:G")))

    If rng Is Nothing Then Exit Sub



    Application.EnableEvents = False

    Application.ScreenUpdating = False

    Application.Calculation = xlCalculationManual



    On Error GoTo Cleanup



    ' --- Find max invoice number ---

    lastRow = Me.Cells(Me.Rows.count, 9).End(xlUp).row  ' Column I = Invoice #

    maxSeq = 0

    For i = START_ROW To lastRow

        If IsNumeric(Me.Cells(i, 9).value) Then

            If CLng(Me.Cells(i, 9).value) > maxSeq Then maxSeq = CLng(Me.Cells(i, 9).value)

        End If

    Next i

    seqNum = maxSeq



    For Each cell In rng

        If cell.row >= START_ROW And Me.Cells(cell.row, 1).value <> "TOTAL" Then

            

            If Me.Cells(cell.row, 6).value <> "" Then  ' F = Amount

                

                ' ---- Apply formulas ----

                Me.Cells(cell.row, 6).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                Me.Cells(cell.row, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                Me.Cells(cell.row, 12).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                Me.Cells(cell.row, 14).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                

                ' Output VAT = (Amount - Discount) * 0.107142857

                Me.Cells(cell.row, 12).Formula = "=IF(F" & cell.row & "-G" & cell.row & "=0, 0, (F" & cell.row & "-G" & cell.row & ")*0.107142857)"

                ' Net Sales = Amount - Discount - Output VAT

                Me.Cells(cell.row, 14).Formula = "=F" & cell.row & "-G" & cell.row & "-L" & cell.row

                

                ' ---- Auto-increment invoice number ----

                If Me.Cells(cell.row, 9).value = "" _

                   And Me.Cells(cell.row, 4).value <> "" _

                   And Me.Cells(cell.row, 6).value <> "" Then

                    seqNum = seqNum + 1

                    Me.Cells(cell.row, 9).value = seqNum

                End If

            End If

        End If

    Next cell



    ' ---- CRJ SYNCHRONIZATION ----

    For Each cell In rng

        If cell.row >= START_ROW _

           And Me.Cells(cell.row, 9).value <> "" _

           And Me.Cells(cell.row, 9).value <> "TOTAL" _

           And Me.Cells(cell.row, 6).value <> "" Then

            Call SyncToCRJ(cell.row)

        End If

    Next cell



Cleanup:

    Application.EnableEvents = True

    Application.ScreenUpdating = True

    Application.Calculation = xlCalculationAutomatic

End Sub









