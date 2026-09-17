'==============================================================================
' modGLAggregation - GL Monthly Aggregation Engine
' Purpose: explicit GL calculation separate from Graphify chart logic
'==============================================================================

Option Explicit

'==============================================================================
' Phase 1: Workbook Column Mapping (from Global-Smile_2026-v7.9.xlsm mapping)
'==============================================================================

' PJ Sheet (sheet1.xml) - Purchases Journal
' Column mapping from PJ_COL_ constants in modEngine.bas
Private Const PJ_COL_COA As Integer = 7       ' Column G - Chart of Accounts
Private Const PJ_COL_NAME As Integer = 5       ' Column E - Supplier Name
Private Const PJ_COL_REF As Integer = 8        ' Column H - Reference
Private Const PJ_COL_GROSS As Integer = 14     ' Column N - Gross Amount
Private Const PJ_COL_NET As Integer = 10       ' Column J - Net Amount
Private Const PJ_COL_VAT As Integer = 13       ' Column M - VAT Amount
Private Const PJ_COL_DATE As Integer = 2       ' Column B - Date (row 2 of sheet)

' CDJ Sheet (sheet2.xml) - Cash Disbursement Journal
' Column mapping from CDJ_COL_ constants in modEngine.bas
Private Const CDJ_COL_DATE As Integer = 3      ' Column C - Date
Private Const CDJ_COL_PARTICULARS As Integer = 5 ' Column E - Particulars/Supplier
Private Const CDJ_COL_CASH As Integer = 6      ' Column F - Cash
Private Const CDJ_COL_VAT As Integer = 7       ' Column G - VAT
Private Const CDJ_COL_WHTAX As Integer = 8     ' Column H - Withholding Tax
Private Const CDJ_COL_DEBIT As Integer = 20    ' Column T - Debit
Private Const CDJ_COL_NOTE As Integer = 24     ' Column X - Note
Private Const CDJ_COL_SUNDRY As Integer = 19   ' Column S - Sundry Account (recently added)
Private Const CDJ_DATA_START As Integer = 15   ' First data row

' SJ Sheet (sheet4.xml) - Sales Journal
Private Const SJ_COL_DATE As Integer = 3       ' Column C - Date
Private Const SJ_COL_CUSTOMER As Integer = 6   ' Column F - Customer
Private Const SJ_COL_INVOICE As Integer = 7    ' Column G - Invoice Number
Private Const SJ_COL_AMOUNT As Integer = 10    ' Column I - Amount
Private Const SJ_COL_DISCOUNT As Integer = 11  ' Column J - Discount
Private Const SJ_COL_OUTPUTVAT As Integer = 13 ' Column O - Output VAT
Private Const SJ_COL_NETSALES As Integer = 15  ' Column P - Net Sales

' CRJ Sheet (sheet5.xml) - Credit Journal
Private Const CRJ_COL_DATE As Integer = 4      ' Column D - Date
Private Const CRJ_COL_CUSTOMER As Integer = 6  ' Column F - Customer
Private Const CRJ_COL_INVOICE As Integer = 7   ' Column G - Invoice
Private Const CRJ_COL_CASH As Integer = 8      ' Column H - Cash
Private Const CRJ_COL_OUTPUTVAT As Integer = 10 ' Column J - Output VAT
Private Const CRJ_COL_EXEMPT As Integer = 11   ' Column K - Exempt
Private Const CRJ_COL_SALES As Integer = 12    ' Column L - Sales
Private Const CRJ_COL_NOTE As Integer = 15     ' Column O - Note

' GJ Sheet (sheet6.xml) - General Journal
Private Const GJ_COL_DATE As Integer = 3       ' Column C - Date
Private Const GJ_COL_CUSTOMER As Integer = 6   ' Column F - Customer
Private Const GJ_COL_INVOICE As Integer = 7    ' Column G - Invoice
Private Const GJ_COL_CASH As Integer = 8       ' Column H - Cash
Private Const GJ_COL_OUTPUTVAT As Integer = 10 ' Column J - Output VAT
Private Const GJ_COL_EXEMPT As Integer = 11    ' Column K - Exempt
Private Const GJ_COL_SALES As Integer = 12     ' Column L - Sales

' SUPPLIERS DATA Sheet (sheet8.xml)
Private Const SUP_COL_TIN As Integer = 3       ' Column B - TIN
Private Const SUP_COL_NAME As Integer = 4      ' Column C - Name
Private Const SUP_COL_ADDRESS As Integer = 5   ' Column D - Address
Private Const SUP_COL_COA As Integer = 6       ' Column E - Chart of Accounts

'==============================================================================
' Phase 2: GL Aggregation Functions
'==============================================================================

'==============================================================================
' BuildGLMonthlyTotals
' Purpose: Build monthly totals for all accounts from posted journals
' Input: targetSheet - CDJ sheet name, startDate, endDate
' Output: 2D array(12, NAccounts) of monthly net movements
'==============================================================================
Public Function BuildGLMonthlyTotals(ByVal targetSheet As String, _
    startDate As Date, endDate As Date) As Variant
    
    ' Initialize 12xN array (months x accounts)
    Dim result(1 To 12, 0 To 0) As Double
    ReDim result(1 To 12, 0 To 0)
    
    ' TODO: Implement journal reading and aggregation
    ' Read journals from PJ, CDJ, SJ, CRJ sheets
    ' Map accounts via COA
    ' Sum debit/credit per month
    ' Return result(month, account) = net movement
    
    BuildGLMonthlyTotals = result
End Function

'==============================================================================
' GetAccountMonthlyDebit
' Purpose: Get total debit for a specific account in a specific month
' Input: accountName - account identifier
'        month - 1-12
'        targetSheet - sheet name
' Output: Double - total debit amount
'==============================================================================
Public Function GetAccountMonthlyDebit(ByVal accountName As String, _
    ByVal month As Integer, ByVal targetSheet As String) As Double
    
    ' TODO: Implement per-account monthly debit extraction
    ' Read from journal sheets, match account via COA
    ' Return sum of debit entries for that account in that month
    
    GetAccountMonthlyDebit = 0
End Function

'==============================================================================
' GetAccountMonthlyCredit
' Purpose: Get total credit for a specific account in a specific month
' Input: accountName - account identifier
'        month - 1-12
'        targetSheet - sheet name
' Output: Double - total credit amount
'==============================================================================
Public Function GetAccountMonthlyCredit(ByVal accountName As String, _
    ByVal month As Integer, ByVal targetSheet As String) As Double
    
    ' TODO: Implement per-account monthly credit extraction
    ' Return sum of credit entries for that account in that month
    
    GetAccountMonthlyCredit = 0
End Function

'==============================================================================
' CalculateEndingBalance
' Purpose: Calculate ending balance for an account through a given month
' Input: accountName - account identifier
'        month - 1-12 (running total through this month)
'        openingBalance - opening balance at month 0
'        targetSheet - sheet name
' Output: Double - ending balance
'==============================================================================
Public Function CalculateEndingBalance(ByVal accountName As String, _
    ByVal month As Integer, ByVal openingBalance As Double, _
    ByVal targetSheet As String) As Double
    
    ' Calculate running balance:
    ' EndingBalance(M) = EndingBalance(M-1) + NetMovement(M)
    ' NetMovement(M) = Debit(M) - Credit(M)
    
    Dim netMovement As Double
    Dim runningBalance As Double
    
    runningBalance = openingBalance
    
    For m = 1 To month
        netMovement = GetAccountMonthlyDebit(accountName, m, targetSheet) - _
                     GetAccountMonthlyCredit(accountName, m, targetSheet)
        runningBalance = runningBalance + netMovement
    Next m
    
    CalculateEndingBalance = runningBalance
End Function

'==============================================================================
' RefreshGL
' Purpose: Refresh all GL balances for the fiscal year
' Input: targetSheet - sheet name, openingBalances(1 To NAccounts)
' Output: 2D array(12, NAccounts) of ending balances
'==============================================================================
Public Function RefreshGL(ByVal targetSheet As String, _
    Optional openingBalances() As Double) As Variant
    
    ' Initialize result array
    Dim result(1 To 12, 0 To 0) As Double
    ReDim result(1 To 12, 0 To 0)
    
    ' TODO: Implement full GL refresh
    ' For each account, calculate ending balance through month 12
    ' Return result(month, account) = ending balance
    
    RefreshGL = result
End Function

'==============================================================================
' ValidateGL
' Purpose: Validate GL consistency - total debits = total credits
' Input: targetSheet - sheet name
' Output: Boolean - True if balanced
'==============================================================================
Public Function ValidateGL(ByVal targetSheet As String) As Boolean
    
    ' TODO: Implement GL validation
    ' Check: Σ Debits = Σ Credits for all posted journals
    ' Check: For each account, debits - credits = net movement consistent
    
    ValidateGL = True ' Placeholder - implement actual validation
End Function

'==============================================================================
' MapPJColumns - Map PJ sheet columns to accounting fields
'==============================================================================
Public Function MapPJColumns() As Object
    
    ' Return mapping of PJ column numbers to accounting field names
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")
    
    mapping.Add "COA", PJ_COL_COA
    mapping.Add "Name", PJ_COL_NAME
    mapping.Add "Reference", PJ_COL_REF
    mapping.Add "GrossAmount", PJ_COL_GROSS
    mapping.Add "NetAmount", PJ_COL_NET
    mapping.Add "VATAmount", PJ_COL_VAT
    mapping.Add "Date", PJ_COL_DATE
    
    MapPJColumns = mapping
End Function

'==============================================================================
' MapCDJColumns - Map CDJ sheet columns to accounting fields
'==============================================================================
Public Function MapCDJColumns() As Object
    
    ' Return mapping of CDJ column numbers to accounting field names
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")
    
    mapping.Add "Date", CDJ_COL_DATE
    mapping.Add "Particulars", CDJ_COL_PARTICULARS
    mapping.Add "Cash", CDJ_COL_CASH
    mapping.Add "VAT", CDJ_COL_VAT
    mapping.Add "WHTAX", CDJ_COL_WHTAX
    mapping.Add "Debit", CDJ_COL_DEBIT
    mapping.Add "Note", CDJ_COL_NOTE
    mapping.Add "Sundry", CDJ_COL_SUNDRY
    
    MapCDJColumns = mapping
End Function

'==============================================================================
' GetAccountFromCOA - Map PJ COA code to account name
'==============================================================================
Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    
    ' TODO: Implement COA code to account name mapping
    ' Could read from SUPPLIERS DATA sheet or a COA lookup table
    ' For now, return the code as-is
    
    GetAccountFromCOA = coaCode
End Function

'==============================================================================
' LoadCOALookup - Load Chart of Accounts from SUPPLIERS DATA sheet
'==============================================================================
Public Sub LoadCOALookup()
    
    ' TODO: Load COA lookup from SUPPLIERS DATA sheet (sheet8.xml)
    ' Column B: TIN, Column C: Name, Column D: Address, Column E: COA
    ' Populate a dictionary for quick lookup
    
End Sub