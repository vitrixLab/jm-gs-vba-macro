Attribute VB_Name = "Sheet9"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Const START_ROW As Long = 8

Private Const COL_TIN As Integer = 4

Private Const COL_NAME As Integer = 5

Private Const COL_ADDRESS As Integer = 6

Private Const COL_COA As Integer = 7

Private Const COL_REF As Integer = 8

Private Const COL_DEBIT As Integer = 10

Private Const COL_DEBIT_MIRROR As Integer = 9

Private Const COL_FULL_ID As Integer = 17



Private gMaxSeq As Long



' ============================================================

' COA List, SelectCOA, GetMaxSequence, GetNextSequence

' (Same as PJ – copy from PJ sheet)

' ============================================================



' ============================================================

' FIND REFERENCE IN CDJ

' ============================================================

Private Function FindReferenceInCDJ(ref As String) As Long

    Dim wsCDJ As Worksheet, found As Range

    On Error Resume Next

    Set wsCDJ = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0

    If wsCDJ Is Nothing Then FindReferenceInCDJ = 0: Exit Function

    

    Set found = wsCDJ.Range("Y:Y").Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)

    If found Is Nothing Then Set found = wsCDJ.Range("Z:Z").Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then FindReferenceInCDJ = found.row Else FindReferenceInCDJ = 0

End Function



' ============================================================

' SYNC NON-VAT TO CDJ

' ============================================================

Private Sub SyncToCDJ_NonVat(sourceRow As Long)

    ' CDJ values must land where the titles say (C=Date, E=Particulars, F=Cash In
    ' Bank, G=Input VAT, H=WHT, N-S=category, T=Debit, X/Y/Z=logs) - the engine
    ' owns that layout now, so delegate instead of writing A-E/M-T here.

    modEngine.SyncPJNonVatToCDJ Me, sourceRow

End Sub




' ============================================================

' HELPER FUNCTIONS (Copy from PJ sheet)

' GetExpenseColumn, CalculateWithholdingTax

' ============================================================



' ============================================================

' MAIN TRIGGER

' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range, cell As Range

    Dim currentTimestamp As String

    

    Set rng = Intersect(Target, Union(Me.Range("D:D"), Me.Range("E:E"), _

                                      Me.Range("F:F"), Me.Range("G:G"), _

                                      Me.Range("I:I")))

    If rng Is Nothing Then Exit Sub



    Application.EnableEvents = False

    Application.ScreenUpdating = False

    Application.Calculation = xlCalculationManual



    On Error GoTo Cleanup

    currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")



    For Each cell In rng

        If cell.row >= START_ROW Then

            

            If cell.Column = COL_TIN Or cell.Column = COL_NAME Then

                Dim tinVal As String, nameVal As String

                tinVal = Trim(Me.Cells(cell.row, COL_TIN).value)

                nameVal = Trim(Me.Cells(cell.row, COL_NAME).value)

                

                If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then

                    Call CalculateTotals(cell.row): GoTo NextCell

                End If

                

                If tinVal <> "" Then

                    Call LookupByTIN_NonVat(cell.row, tinVal)

                ElseIf nameVal <> "" Then

                    Call LookupByName_NonVat(cell.row, nameVal)

                Else

                    Call ClearSupplierRow(cell.row)

                End If

            End If

            

            ' ---- NON-VAT AMOUNT (No VAT) ----

            If Me.Cells(cell.row, "I").value <> "" _

               And Me.Cells(cell.row, COL_TIN).value <> "TOTAL" _

               And Me.Cells(cell.row, COL_NAME).value <> "TOTAL" Then

                

                Me.Cells(cell.row, COL_DEBIT).value = Me.Cells(cell.row, "I").value

                Me.Cells(cell.row, COL_DEBIT_MIRROR).value = Me.Cells(cell.row, "I").value

                Me.Cells(cell.row, 13).value = 0  ' Credit = 0

                Me.Cells(cell.row, 12).value = 0  ' Credit Mirror = 0

                

                If Me.Cells(cell.row, COL_TIN).value <> "" _

                   And Me.Cells(cell.row, COL_NAME).value <> "" _

                   And Me.Cells(cell.row, COL_COA).value <> "" _

                   And Me.Cells(cell.row, COL_REF).value = "" Then

                    

                    Dim seqNum As Long

                    seqNum = GetNextSequence()

                    Me.Cells(cell.row, COL_REF).value = "PNV-" & Format(seqNum, "00000")

                    Me.Cells(cell.row, COL_FULL_ID).value = "PNVINV-" & currentTimestamp & "-" & Format(seqNum, "00000")

                End If

            End If

        End If

NextCell:

    Next cell



    ' ---- CDJ SYNCHRONIZATION ----

    For Each cell In rng

        If cell.row >= START_ROW _

           And Me.Cells(cell.row, COL_REF).value <> "" _

           And Me.Cells(cell.row, COL_REF).value <> "TOTAL" Then

            Call SyncToCDJ_NonVat(cell.row)

        End If

    Next cell



Cleanup:

    Application.EnableEvents = True

    Application.ScreenUpdating = True

    Application.Calculation = xlCalculationAutomatic

End Sub



' ============================================================

' LOOKUP FUNCTIONS (Copy from PJ sheet – rename to _NonVat)

' ============================================================









