'==============================================================================
' modGLAggregation - GL Monthly Aggregation Engine (Implementation)
' Purpose: explicit GL calculation separate from Graphify chart logic
'==============================================================================

Option Explicit

'==============================================================================
' Phase 1: Workbook Column Mapping (from Global-Smile_2026-v7.9.xlsm mapping)
'==============================================================================

' PJ Sheet (sheet1.xml) - Purchases Journal
Private Const PJ_COL_COA As Integer = 7       ' Column G - Chart of Accounts
Private Const PJ_COL_NAME As Integer = 5       ' Column E - Supplier Name
Private Const PJ_COL_REF As Integer = 8        ' Column H - Reference
Private Const PJ_COL_GROSS As Integer = 14     ' Column N - Gross Amount
Private Const PJ_COL_NET As Integer = 10       ' Column J - Net Amount
Private Const PJ_COL_VAT As Integer = 13       ' Column M - VAT Amount
Private Const PJ_COL_DATE As Integer = 2       ' Column B - Date (row 2 of sheet)

' CDJ Sheet (sheet2.xml) - Cash Disbursement Journal
Private Const CDJ_COL_DATE As Integer = 3      ' Column C - Date
Private Const CDJ_COL_PARTICULARS As Integer = 5 ' Column E - Particulars/Supplier
Private Const CDJ_COL_CASH As Integer = 6      ' Column F - Cash
Private Const CDJ_COL_VAT As Integer = 7       ' Column G - VAT
Private Const CDJ_COL_WHTAX As Integer = 8     ' Column H - Withholding Tax
Private Const CDJ_COL_DEBIT As Integer = 20    ' Column T - Debit
Private Const CDJ_COL_NOTE As Integer = 24     ' Column X - Note
Private Const CDJ_COL_SUNDRY As Integer = 19   ' Column S - Sundry Account
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
' Collection Objects for COA Lookup
'==============================================================================
Private pCOADict As Object

'==============================================================================
' Phase 2: Initialize COA Lookup
'==============================================================================
Public Sub InitializeCOALookup()
    
    Set pCOADict = CreateObject("Scripting.Dictionary")
    
    ' Load COA from SUPPLIERS DATA sheet
    ' Column E (SUP_COL_COA) contains COA codes
    ' Column C (SUP_COL_NAME) contains account names
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, SUP_COL_TIN).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim tin As String, name As String, coa As String
        tin = Trim(ws.Cells(i, SUP_COL_TIN).value)
        name = Trim(ws.Cells(i, SUP_COL_NAME).value)
        coa = Trim(ws.Cells(i, SUP_COL_COA).value)
        
        If tin <> "" And coa <> "" Then
            pCOADict(tin) = Array(name, coa)
        If name <> "" And coa <> "" Then
            pCOADict(name) = Array(tin, coa)
        End If
    Next i
    
End Sub

'==============================================================================
' Phase 2: GetAccountFromCOA - Map COA code to account name
'==============================================================================
Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    
    If pCOADict Is Nothing Then InitializeCOALookup
    
    Dim key As Variant
    For Each key In pCOADict.Keys
        If pCOADict(key)(1) = coaCode Then
            GetAccountFromCOA = pCOADict(key)(0)
            Exit Function
        End If
    Next key
    
    ' Fallback: return the code itself
    GetAccountFromCOA = coaCode
End Function

'==============================================================================
' Phase 2: GetAccountNameFromTIN - Map TIN to account name
'==============================================================================
Public Function GetAccountNameFromTIN(ByVal tin As String) As String
    
    If pCOADict Is Nothing Then InitializeCOALookup
    
    If pCOADict.Exists(tin) Then
        GetAccountNameFromTIN = pCOADict(key)(0)
    Else
        GetAccountNameFromTIN = tin
    End If
    
End Function

'==============================================================================
' Phase 2: GetMonthlyJournalData - Extract journal data for a month
'==============================================================================
Public Function GetMonthlyJournalData(ByVal sheetName As String, _
    ByVal month As Integer, ByVal year As Integer) As Variant
    
    ' Extract journal rows for a specific month/year
    ' Returns array of (date, account, debit, credit, description)
    
    Dim result() As Variant
    ReDim result(0 To 0, 0 To 3) As Double
    result(1, 1) = 0: result(1, 2) = 0: result(1, 3) = 0 ' placeholder
    
    ' TODO: Implement based on sheetName
    ' - Read from CDJ, SJ, CRJ, or GJ sheet
    ' - Filter by month/year
    ' - Return array: (date, accountCode, debit, credit)
    
    GetMonthlyJournalData = result
End Function

'==============================================================================
' Phase 2: CalculateMonthlyNet - Calculate net movement for an account in a month
'==============================================================================
Public Function CalculateMonthlyNet(ByVal accountCode As String, _
    ByVal month As Integer, ByVal year As Integer, _
    ByVal sheetName As String) As Double
    
    ' Calculate: NetMovement = Debit - Credit for account in month/year
    ' Uses journal data extraction and COA mapping
    
    Dim transactions() As Variant
    transactions = GetMonthlyJournalData(sheetName, month, year)
    
    Dim totalDebit As Double, totalCredit As Double
    totalDebit = 0: totalCredit = 0
    
    ' TODO: Filter by accountCode and sum debits/credits
    ' For now return 0
    
    CalculateMonthlyNet = 0
End Function

'==============================================================================
' Phase 3: ValidateGLConsistency - Check GL balance integrity
'==============================================================================
Public Function ValidateGLConsistency(ByVal targetSheet As String) As Boolean
    
    ' Check: For the given sheet, total debits = total credits
    ' Also check: ending balances are consistent with beginning balances + net movements
    
    ValidateGLConsistency = True ' Placeholder - implement actual validation
    
    ' TODO: Implement
    ' 1. Read all debit/credit entries from the sheet
    ' 2. Sum total debits and total credits
    ' 3. If |TotalDebits - TotalCredits| > tolerance, set False
    ' 4. Check ending balances consistency
    
End Function

'==============================================================================
' Phase 2: BuildGLOpeningBalance - Get opening balance for January
'==============================================================================
Public Function BuildGLOpeningBalance(ByVal accountName As String, _
    ByVal sheetName As String) As Double
    
    ' Get the opening balance (January beginning balance) for an account
    ' This would typically come from prior year data or be 0 for new accounts
    
    ' TODO: Implement - could read from opening balance sheet or be 0
    BuildGLOpeningBalance = 0
    
End Function

'==============================================================================
' Phase 2: CalculateEndingBalanceFull - Full year running balance
'==============================================================================
Public Function CalculateEndingBalanceFull(ByVal accountCode As String, _
    ByVal year As Integer, ByVal startBalance As Double, _
    ByVal targetSheet As String) As Variant
    
    ' Calculate ending balance for all 12 months
    ' Input: startBalance = January opening balance
    ' Output: array(1 To 12) of ending balances
    
    Dim balances(1 To 12) As Double
    Dim m As Integer
    
    balances(1) = startBalance + CalculateMonthlyNet(accountCode, 1, year, targetSheet)
    
    For m = 2 To 12
        balances(m) = balances(m - 1) + CalculateMonthlyNet(accountCode, m, year, targetSheet)
    Next m
    
    CalculateEndingBalanceFull = balances
    
End Function

'==============================================================================
' Phase 2: RefreshAllGL - Refresh GL for all accounts in the workbook
'==============================================================================
Public Sub RefreshAllGL(ByVal targetSheet As String, _
    ByVal startBalances() As Double)
    
    ' Refresh GL balances for all accounts through the fiscal year
    ' Input: startBalances(1 To NAccounts) = January opening balances
    ' Output: writes ending balances to a results sheet or collection
    
    ' TODO: Implement full refresh
    ' 1. For each account, calculate ending balances for months 1-12
    ' 2. Write results to appropriate location
    ' 3. Optionally validate consistency
    
End Sub

'==============================================================================
' Phase 1: MapSheetColumns - Return column mapping for a sheet
'==============================================================================
Public Function MapSheetColumns(ByVal sheetName As String) As Object
    
    ' Return a dictionary mapping field names to column numbers
    ' for the specified sheet name
    
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")
    
    Select Case sheetName
        Case "PJ"
            mapping.Add "COA", PJ_COL_COA
            mapping.Add "Name", PJ_COL_NAME
            mapping.Add "Reference", PJ_COL_REF
            mapping.Add "Gross", PJ_COL_GROSS
            mapping.Add "Net", PJ_COL_NET
            mapping.Add "VAT", PJ_COL_VAT
            mapping.Add "Date", PJ_COL_DATE
            
        Case "CDJ"
            mapping.Add "Date", CDJ_COL_DATE
            mapping.Add "Particulars", CDJ_COL_PARTICULARS
            mapping.Add "Cash", CDJ_COL_CASH
            mapping.Add "VAT", CDJ_COL_VAT
            mapping.Add "WHTAX", CDJ_COL_WHTAX
            mapping.Add "Debit", CDJ_COL_DEBIT
            mapping.Add "Note", CDJ_COL_NOTE
            mapping.Add "Sundry", CDJ_COL_SUNDRY
            
        Case "SJ"
            mapping.Add "Date", SJ_COL_DATE
            mapping.Add "Customer", SJ_COL_CUSTOMER
            mapping.Add "Invoice", SJ_COL_INVOICE
            mapping.Add "Amount", SJ_COL_AMOUNT
            mapping.Add "Discount", SJ_COL_DISCOUNT
            mapping.Add "OutputVAT", SJ_COL_OUTPUTVAT
            mapping.Add "NetSales", SJ_COL_NETSALES
            
        Case "CRJ"
            mapping.Add "Date", CRJ_COL_DATE
            mapping.Add "Customer", CRJ_COL_CUSTOMER
            mapping.Add "Invoice", CRJ_COL_INVOICE
            mapping.Add "Cash", CRJ_COL_CASH
            mapping.Add "OutputVAT", CRJ_COL_OUTPUTVAT
            mapping.Add "Exempt", CRJ_COL_EXEMPT
            mapping.Add "Sales", CRJ_COL_SALES
            
        Case "GJ"
            mapping.Add "Date", GJ_COL_DATE
            mapping.Add "Customer", GJ_COL_CUSTOMER
            mapping.Add "Invoice", GJ_COL_INVOICE
            mapping.Add "Cash", GJ_COL_CASH
            mapping.Add "OutputVAT", GJ_COL_OUTPUTVAT
            mapping.Add "Exempt", GJ_COL_EXEMPT
            mapping.Add "Sales", GJ_COL_SALES
    End Select
    
    MapSheetColumns = mapping
    
End Function

'==============================================================================
' Utility: GetFinancialMonth - Convert date to month number (1-12)
'==============================================================================
Public Function GetFinancialMonth(ByVal dt As Date) As Integer
    
    GetFinancialMonth = Month(dt)
    
End Function

'==============================================================================
' Utility: GetFinancialYear - Get fiscal year from date
'==============================================================================
Public Function GetFinancialYear(ByVal dt As Date) As Integer
    
    ' If month >= 7, fiscal year ends in current year
    ' If month < 7, fiscal year ends in previous year
    If Month(dt) >= 7 Then
        GetFinancialYear = Year(dt)
    Else
        GetFinancialYear = Year(dt) - 1
    End If
    
End Function

End Module