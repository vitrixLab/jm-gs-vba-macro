Attribute VB_Name = "Sheet3"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Const START_ROW As Long = 10

Private Const CRJ_SHEET_NAME As String = "CRJ"



' SJ Columns (now final)

Private Const SJ_COL_DATE As Integer = 3

Private Const SJ_COL_CUSTOMER As Integer = 6

Private Const SJ_COL_CUSTOMER_ADDRESS As Integer = 7

Private Const SJ_COL_DESCRIPTION As Integer = 8

Private Const SJ_COL_AMOUNT As Integer = 10

Private Const SJ_COL_DISCOUNT As Integer = 11

Private Const SJ_COL_OUTPUTVAT_MIRROR As Integer = 12

Private Const SJ_COL_OUTPUTVAT As Integer = 13

Private Const SJ_COL_NETSALES_MIRROR As Integer = 14

Private Const SJ_COL_NETSALES As Integer = 15

Private Const SJ_COL_INVOICE As Integer = 21         ' U – auto short number (SJ-NNNNN)

Private Const SJ_COL_FULLID As Integer = 22          ' V – auto full ID (SJINV-…)

Private Const SJ_COL_PHYSICAL_INV As Integer = 9     ' I – manual physical invoice (NEVER overwritten)



Private Sub Worksheet_Activate()

    On Error Resume Next

    modEngine.InitializeCRJ

    On Error GoTo 0

End Sub



Private Sub LookupCustomer(rowNum As Long, name As String)

    name = Trim(name)

    If name = "" Then Exit Sub

    Me.Cells(rowNum, SJ_COL_CUSTOMER_ADDRESS).value = "NCR Metro Manila"

    Me.Cells(rowNum, SJ_COL_DESCRIPTION).value = "Full Payment for Service"

End Sub



Private Sub ClearCustomerRow(rowNum As Long)

    Me.Cells(rowNum, SJ_COL_CUSTOMER_ADDRESS).value = ""

    Me.Cells(rowNum, SJ_COL_DESCRIPTION).value = ""

End Sub



Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rngLookup As Range, rngSync As Range, cell As Range

    Dim currentTimestamp As String, newInvoice As Variant



    ' ---- PART 1: LOOKUP, FORMULAS, INVOICE STAMPING ----

    Set rngLookup = Intersect(Target, Union(Me.Range("F:F"), Me.Range("G:G"), Me.Range("J:J"), Me.Range("K:K")))

    If Not rngLookup Is Nothing Then

        Application.EnableEvents = False

        Application.ScreenUpdating = False

        Application.Calculation = xlCalculationManual

        On Error GoTo CleanupLookup

        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")



        For Each cell In rngLookup

            If cell.row >= START_ROW Then



                ' ---- SKIP TOTAL ROW ----

                If UCase(Trim(Me.Cells(cell.row, 5).value)) = "TOTAL" Then GoTo NextLookup



                If cell.Column = SJ_COL_CUSTOMER Then

                    Call LookupCustomer(cell.row, Trim(Me.Cells(cell.row, SJ_COL_CUSTOMER).value))

                End If



                If Me.Cells(cell.row, SJ_COL_AMOUNT).value <> "" And _

                   Me.Cells(cell.row, SJ_COL_AMOUNT).value <> "TOTAL" Then



                    ' Formats

                    Me.Cells(cell.row, SJ_COL_AMOUNT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                    Me.Cells(cell.row, SJ_COL_DISCOUNT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                    Me.Cells(cell.row, SJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                    Me.Cells(cell.row, SJ_COL_NETSALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"



                    ' Formulas

                    Me.Cells(cell.row, SJ_COL_OUTPUTVAT).Formula = "=IF(J" & cell.row & "-K" & cell.row & "=0, 0, (J" & cell.row & "-K" & cell.row & ")*0.107142857)"

                    Me.Cells(cell.row, SJ_COL_NETSALES).Formula = "=J" & cell.row & "-K" & cell.row & "-M" & cell.row

                    Me.Cells(cell.row, SJ_COL_OUTPUTVAT_MIRROR).Formula = "=M" & cell.row

                    Me.Cells(cell.row, SJ_COL_NETSALES_MIRROR).Formula = "=O" & cell.row



                    ' ---- STAMP AUTO INVOICE (U) & FULL ID (V) ----

                    If Me.Cells(cell.row, SJ_COL_INVOICE).value = "" Then

                        newInvoice = modEngine.GetNextInvoiceNumber()   ' e.g. 304

                        Me.Cells(cell.row, SJ_COL_INVOICE).value = "SJ-" & Format(newInvoice, "00000")



                        ' Read the manual physical invoice from column I

                        Dim physicalInv As String

                        If IsNumeric(Me.Cells(cell.row, SJ_COL_PHYSICAL_INV).value) Then

                            physicalInv = Format(CLng(Me.Cells(cell.row, SJ_COL_PHYSICAL_INV).value), "00000")

                        Else

                            physicalInv = "00000"   ' fallback if I is blank

                        End If



                        Me.Cells(cell.row, SJ_COL_FULLID).value = "SJINV-" & currentTimestamp & "-" & physicalInv & "-" & Format(newInvoice, "00000")



                        If SheetExists(CRJ_SHEET_NAME) Then

                            Call modEngine.SyncSJToCRJ(Me, cell.row)

                        End If

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



    ' ---- PART 2: SYNC (Triggered by change in U – Invoice column) ----

    Set rngSync = Intersect(Target, Me.Range("U:U"))

    If Not rngSync Is Nothing Then

        Application.EnableEvents = False

        Application.ScreenUpdating = False

        Application.Calculation = xlCalculationManual

        On Error GoTo CleanupSync



        For Each cell In rngSync

            If cell.row >= START_ROW Then

                ' ---- SKIP TOTAL ROW ----

                If UCase(Trim(Me.Cells(cell.row, 5).value)) = "TOTAL" Then GoTo NextSync



                If Me.Cells(cell.row, SJ_COL_INVOICE).value <> "" And _

                   Me.Cells(cell.row, SJ_COL_INVOICE).value <> "TOTAL" Then

                    If SheetExists(CRJ_SHEET_NAME) Then

                        Call modEngine.SyncSJToCRJ(Me, cell.row)

                    End If

                End If

            End If

NextSync:

        Next cell



CleanupSync:

        Application.EnableEvents = True

        Application.ScreenUpdating = True

        Application.Calculation = xlCalculationAutomatic

    End If

End Sub



Sub ManualSyncToCRJ()

    Dim rowNum As Long

    rowNum = ActiveCell.row

    If rowNum < START_ROW Then

        MsgBox "Select a row below row 10.", vbExclamation

        Exit Sub

    End If



    ' ---- SKIP TOTAL ROW ----

    If UCase(Trim(Me.Cells(rowNum, 5).value)) = "TOTAL" Then

        MsgBox "Cannot sync the TOTAL row.", vbExclamation

        Exit Sub

    End If



    If Me.Cells(rowNum, SJ_COL_INVOICE).value = "" Then

        MsgBox "No invoice number in column U. Enter Amount first.", vbExclamation

        Exit Sub

    End If



    If Not SheetExists(CRJ_SHEET_NAME) Then

        MsgBox "CRJ sheet not found. Cannot sync.", vbCritical

        Exit Sub

    End If



    On Error Resume Next

    Call modEngine.SyncSJToCRJ(Me, rowNum)

    If Err.Number = 0 Then

        MsgBox "Successfully synced row " & rowNum & " to CRJ.", vbInformation

    Else

        MsgBox "Sync failed: " & Err.description, vbExclamation

    End If

    On Error GoTo 0

End Sub



Private Function SheetExists(sheetName As String) As Boolean

    On Error Resume Next

    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing

    On Error GoTo 0

End Function





