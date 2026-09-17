Attribute VB_Name = "Sheet12"

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

Private Const COL_DEBIT_MIRROR As Integer = 9

Private Const COL_DEBIT As Integer = 10

Private Const COL_CREDIT_MIRROR As Integer = 12

Private Const COL_CREDIT As Integer = 13

Private Const COL_GROSS As Integer = 14

Private Const COL_FULL_ID As Integer = 17



' ============================================================

' COA LIST & SELECT (used only when adding new supplier)

' ============================================================

Private Function GetCOAList() As Variant

    Static coaArray As Variant

    If IsEmpty(coaArray) Then

        coaArray = Array( _

            "Representation", "Clinicians Fee - Corporators", "Value Added Tax Payable", _

            "Pantry Supplies", "Taxes and Licences", "Miscellaneous", "Advances to Employee", _

            "Government ER Share Cont", "Delivery", "Repairs and Maintenance", _

            "Clinicians Fee - Visiting", "Association Dues", "Advertisement", _

            "Clinic Materials and Supplies", "Transportation and Travel", "Fuel and Oil", _

            "Supplies", "Rental", "Communication", "Professional Fees" _

        )

    End If

    GetCOAList = coaArray

End Function



Private Function SelectCOA() As String

    Dim coaList As Variant, i As Integer, inputStr As String, userInput As String, idx As Integer

    coaList = GetCOAList()

    inputStr = "Select Chart of Account:" & vbCrLf

    For i = LBound(coaList) To UBound(coaList)

        inputStr = inputStr & (i + 1) & ". " & coaList(i) & vbCrLf

    Next i

    inputStr = inputStr & "Enter number (1-" & UBound(coaList) + 1 & "):"

    userInput = InputBox(inputStr, "Chart of Account Selection")

    If userInput = "" Then SelectCOA = "": Exit Function

    If IsNumeric(userInput) Then

        idx = CInt(userInput) - 1

        If idx >= LBound(coaList) And idx <= UBound(coaList) Then SelectCOA = coaList(idx): Exit Function

    End If

    MsgBox "Invalid selection.", vbExclamation

    SelectCOA = SelectCOA()

End Function



' ============================================================

' LOOKUP (uses modEngine)

' ============================================================

Private Sub LookupByTIN(rowNum As Long, tin As String)

    Dim data As Variant

    tin = Trim(tin): If tin = "" Then Exit Sub

    data = modEngine.GetSupplierByTIN(tin)

    If IsArray(data) Then

        Me.Cells(rowNum, COL_NAME).value = data(0)

        Me.Cells(rowNum, COL_ADDRESS).value = data(1)

        Me.Cells(rowNum, COL_COA).value = data(2)

    Else

        If MsgBox("TIN '" & tin & "' not found. Add?", vbYesNo + vbQuestion) = vbYes Then

            Dim name As String, address As String, coa As String

            name = Trim(Me.Cells(rowNum, COL_NAME).value): If name = "" Then name = InputBox("Supplier Name:")

            If name = "" Then Exit Sub

            address = Trim(Me.Cells(rowNum, COL_ADDRESS).value): If address = "" Then address = InputBox("Address:")

            coa = Trim(Me.Cells(rowNum, COL_COA).value): If coa = "" Then coa = SelectCOA()

            If coa = "" Then Exit Sub

            Call AddSupplierToMaster(tin, name, coa, address)

            Call LookupByTIN(rowNum, tin)

        End If

    End If

End Sub



Private Sub LookupByName(rowNum As Long, name As String)

    Dim data As Variant

    name = Trim(name): If name = "" Then Exit Sub

    data = modEngine.GetSupplierByName(name)

    If IsArray(data) Then

        Me.Cells(rowNum, COL_TIN).value = data(0)

        Me.Cells(rowNum, COL_ADDRESS).value = data(1)

        Me.Cells(rowNum, COL_COA).value = data(2)

    Else

        If MsgBox("Supplier '" & name & "' not found. Add?", vbYesNo + vbQuestion) = vbYes Then

            Dim tin As String, address As String, coa As String

            tin = Trim(Me.Cells(rowNum, COL_TIN).value): If tin = "" Then tin = InputBox("TIN:")

            If tin = "" Then Exit Sub

            address = Trim(Me.Cells(rowNum, COL_ADDRESS).value): If address = "" Then address = InputBox("Address:")

            coa = Trim(Me.Cells(rowNum, COL_COA).value): If coa = "" Then coa = SelectCOA()

            If coa = "" Then Exit Sub

            Call AddSupplierToMaster(tin, name, coa, address)

            Call LookupByName(rowNum, name)

        End If

    End If

End Sub



' ============================================================

' ADD SUPPLIER TO MASTER + UPDATE CACHE

' ============================================================

Private Sub AddSupplierToMaster(tin As String, name As String, coa As String, address As String)

    Dim ws As Worksheet, lastRow As Long, nextNo As Long, i As Long

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).row

    For i = 2 To lastRow

        If ws.Cells(i, 3).value = tin Then

            MsgBox "TIN already exists!", vbExclamation: Exit Sub

        End If

    Next i



    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row

    If lastRow < 2 Then lastRow = 2

    nextNo = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2))) + 1

    If IsError(nextNo) Then nextNo = 1



    Dim newRow As Long: newRow = lastRow + 1

    With ws

        .Cells(newRow, 2).value = nextNo

        .Cells(newRow, 3).value = tin

        .Cells(newRow, 4).value = UCase(name)

        .Cells(newRow, 5).value = address

        .Cells(newRow, 6).value = coa

    End With

    Call modEngine.AddSupplierToCache(tin, name, address, coa)

    MsgBox "Supplier added.", vbInformation

End Sub



' ============================================================

' HELPERS (Totals, Block Start)

' ============================================================

Private Sub ClearSupplierRow(rowNum As Long)

    Me.Cells(rowNum, COL_NAME).value = ""

    Me.Cells(rowNum, COL_ADDRESS).value = ""

    Me.Cells(rowNum, COL_COA).value = ""

End Sub



Private Sub CalculateTotals(rowNum As Long)

    Dim startRow As Long, endRow As Long, debitSum As Double, creditSum As Double

    startRow = FindBlockStart(rowNum): endRow = rowNum - 1

    If endRow < startRow Then Exit Sub

    debitSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_DEBIT), Me.Cells(endRow, COL_DEBIT)))

    creditSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_CREDIT), Me.Cells(endRow, COL_CREDIT)))

    Me.Cells(rowNum, COL_DEBIT).value = debitSum

    Me.Cells(rowNum, COL_CREDIT).value = creditSum

    Me.Cells(rowNum, COL_DEBIT_MIRROR).value = debitSum

    Me.Cells(rowNum, COL_CREDIT_MIRROR).value = creditSum

End Sub



Private Function FindBlockStart(rowNum As Long) As Long

    Dim i As Long, val As String

    For i = rowNum - 1 To START_ROW Step -1

        val = Trim(Me.Cells(i, COL_TIN).value)

        If LCase(val) = "total" Then FindBlockStart = i + 1: Exit Function

    Next i

    FindBlockStart = START_ROW

End Function



' ============================================================

' MANUAL SYNC MACRO (optional)

' ============================================================

Sub ManualSyncToCDJ()

    Dim rowNum As Long

    rowNum = ActiveCell.row

    If rowNum < START_ROW Then

        MsgBox "Select a row below row 8.", vbExclamation

        Exit Sub

    End If

    If Me.Cells(rowNum, COL_REF).value = "" Then

        MsgBox "No reference ID (PJ-XXXXX) in column H. Enter Gross Amount first.", vbExclamation

        Exit Sub

    End If

    Call modEngine.SyncPJToCDJ(Me, rowNum)

    MsgBox "Sync attempted for row " & rowNum & ". Check target sheet.", vbInformation

End Sub



' ============================================================

' MAIN TRIGGER

' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rngLookup As Range, rngSync As Range, cell As Range

    Dim currentTimestamp As String, newRef As String



    ' ---- PART 1: LOOKUP, VAT, ID STAMPING ----

    Set rngLookup = Intersect(Target, Union(Me.Range("D:D"), Me.Range("E:E"), _

                                            Me.Range("F:F"), Me.Range("G:G"), _

                                            Me.Range("N:N")))

    If Not rngLookup Is Nothing Then

        Application.EnableEvents = False

        Application.ScreenUpdating = False

        Application.Calculation = xlCalculationManual

        On Error GoTo CleanupLookup

        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")



        For Each cell In rngLookup

            If cell.row >= START_ROW Then

                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then

                    Dim tinVal As String, nameVal As String

                    tinVal = Trim(Me.Cells(cell.row, COL_TIN).value)

                    nameVal = Trim(Me.Cells(cell.row, COL_NAME).value)

                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then

                        Call CalculateTotals(cell.row)

                        GoTo NextLookup

                    End If

                    If tinVal <> "" Then

                        Call LookupByTIN(cell.row, tinVal)

                    ElseIf nameVal <> "" Then

                        Call LookupByName(cell.row, nameVal)

                    Else

                        Call ClearSupplierRow(cell.row)

                    End If

                End If



                If Me.Cells(cell.row, COL_GROSS).value <> "" And _

                   Me.Cells(cell.row, COL_TIN).value <> "TOTAL" And _

                   Me.Cells(cell.row, COL_NAME).value <> "TOTAL" Then

                    Me.Cells(cell.row, COL_DEBIT).Formula = "=N" & cell.row & "/1.12"

                    Me.Cells(cell.row, COL_CREDIT).Formula = "=N" & cell.row & "-J" & cell.row

                    Me.Cells(cell.row, COL_DEBIT_MIRROR).Formula = "=J" & cell.row

                    Me.Cells(cell.row, COL_CREDIT_MIRROR).Formula = "=M" & cell.row



                    ' ---- ID STAMPING ----

                    If Me.Cells(cell.row, COL_TIN).value <> "" And _

                       Me.Cells(cell.row, COL_NAME).value <> "" And _

                       Me.Cells(cell.row, COL_COA).value <> "" And _

                       Me.Cells(cell.row, COL_REF).value = "" Then

                        newRef = modEngine.GetNextPJNumber()

                        Me.Cells(cell.row, COL_REF).value = newRef

                        Me.Cells(cell.row, COL_FULL_ID).value = "PJINV-" & currentTimestamp & "-" & Right(newRef, 5)



                        ' ====== FIX: SYNC IMMEDIATELY ======

                        Call modEngine.SyncPJToCDJ(Me, cell.row)

                    End If

                End If

            End If

NextLookup:

        Next cell



CleanupLookup:

        Application.EnableEvents = True

        Application.ScreenUpdating = True

        Application.Calculation = xlCalculationAutomatic

    End If



    ' ---- PART 2: SYNC (Triggered by Column H) ----

    ' Kept as a safety fallback (will do nothing if already synced)

    Set rngSync = Intersect(Target, Me.Range("H:H"))

    If Not rngSync Is Nothing Then

        Application.EnableEvents = False

        Application.ScreenUpdating = False

        Application.Calculation = xlCalculationManual

        On Error GoTo CleanupSync



        For Each cell In rngSync

            If cell.row >= START_ROW Then

                newRef = Trim(Me.Cells(cell.row, COL_REF).value)

                If newRef <> "" And newRef <> "TOTAL" Then

                    If Me.Cells(cell.row, COL_TIN).value <> "" And _

                       Me.Cells(cell.row, COL_NAME).value <> "" And _

                       Me.Cells(cell.row, COL_COA).value <> "" And _

                       Me.Cells(cell.row, COL_GROSS).value <> "" Then

                        Call modEngine.SyncPJToCDJ(Me, cell.row)

                    End If

                End If

            End If

        Next cell



CleanupSync:

        Application.EnableEvents = True

        Application.ScreenUpdating = True

        Application.Calculation = xlCalculationAutomatic

    End If

End Sub





