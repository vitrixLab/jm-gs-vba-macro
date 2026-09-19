Attribute VB_Name = "modGLAggregation"
Option Explicit

' v8.0 deterministic GL engine for the actual workbook layout.
' Posting sources are CDJ, CRJ and GJ. PJ/PJ Non-Vat/SJ are feeder journals
' and are not posted a second time.

Private Function Put(ByVal d As Object, ByVal acct As String, ByVal m As Long, ByVal db As Double, ByVal cr As Double)
    Dim k As String, a As Variant
    k=V8_Norm(acct) & "|" & m
    If Not d.Exists(k) Then d.Add k,Array(0#,0#)
    a=d(k): a(0)=a(0)+db: a(1)=a(1)+cr: d(k)=a
End Function

Private Function LogDate(ByVal s As String) As Date
    Dim p As Long: p=InStrRev(s," | "): If p>0 Then s=Mid$(s,p+3)
    If IsDate(s) Then LogDate=CDate(s)
End Function

Private Function Matrix(ByVal yr As Long,ByRef gd As Double,ByRef gc As Double,ByRef um As Double,ByRef inv As Long,ByRef detail As String) As Object
    Dim d As Object,ws As Worksheet,r As Long,c As Long,m As Long,cur As Long,a As String,raw As String,v As Double,ok As Boolean,dt As Date
    Set d=CreateObject("Scripting.Dictionary"): d.CompareMode=vbTextCompare

    ' CDJ: F:S are signed account postings; positive=debit, negative=credit.
    Set ws=ThisWorkbook.Worksheets("CDJ"): cur=0
    For r=15 To ws.Cells(ws.Rows.Count,3).End(xlUp).Row
        m=V8_Month(ws.Cells(r,3).Value2): If m>0 Then cur=m
        If cur>0 And Len(Trim$(CStr(ws.Cells(r,5).Value2)))>0 And V8_Norm(ws.Cells(r,5).Value2)<>"TOTAL" Then
            For c=6 To 19
                ok=True:v=V8_Number(ws.Cells(r,c).Value2,ok):If Not ok Then inv=inv+1
                If ok And Abs(v)>TOLERANCE Then
                    a=V8_CDJMap(c)
                    If Len(a)=0 Then um=um+Abs(v):detail=detail&"CDJ!"&ws.Cells(r,c).Address(False,False)&" ambiguous mapping"&vbCrLf
                    ElseIf v>=0 Then Put d,a,cur,v,0#:gd=gd+v Else Put d,a,cur,0#,-v:gc=gc-v
                End If
            Next c
        End If
    Next r

    ' CRJ: H is debit; I:L are credit accounts.
    Set ws=ThisWorkbook.Worksheets("CRJ")
    For r=10 To ws.Cells(ws.Rows.Count,7).End(xlUp).Row
        If Len(Trim$(CStr(ws.Cells(r,7).Value2)))>0 And V8_Norm(ws.Cells(r,6).Value2)<>"TOTAL" Then
            dt=0: If IsDate(ws.Cells(r,15).Value2) Then dt=CDate(ws.Cells(r,15).Value2) Else dt=0
            If dt=0 Then
                Dim logText As String, p As Long
                logText=CStr(ws.Cells(r,15).Value2):p=InStrRev(logText," | "):If p>0 Then logText=Mid$(logText,p+3)
                If IsDate(logText) Then dt=CDate(logText)
            End If
            If dt=0 Then inv=inv+1 ElseIf Year(dt)=yr Then
                m=Month(dt):ok=True:v=V8_Number(ws.Cells(r,8).Value2,ok):If Not ok Then inv=inv+1
                If ok And Abs(v)>TOLERANCE Then Put d,"Cash in Bank",m,v,0#:gd=gd+v
                For c=9 To 12
                    ok=True:v=V8_Number(ws.Cells(r,c).Value2,ok):If Not ok Then inv=inv+1
                    If ok And Abs(v)>TOLERANCE Then
                        a=V8_CRJMap(c)
                        If Len(a)=0 Then um=um+Abs(v):detail=detail&"CRJ!"&ws.Cells(r,c).Address(False,False)&" unmapped credit"&vbCrLf Else Put d,a,m,0#,v:gc=gc+v
                    End If
                Next c
            End If
        End If
    Next r

    ' GJ: month marker persists; F=debit, G=credit.
    Set ws=ThisWorkbook.Worksheets("GJ"):cur=0
    For r=11 To ws.Cells(ws.Rows.Count,5).End(xlUp).Row
        m=V8_Month(ws.Cells(r,2).Value2):If m>0 Then cur=m
        raw=Trim$(CStr(ws.Cells(r,5).Value2))
        If cur>0 And Len(raw)>0 And V8_Norm(raw)<>"RECORDING DEPRECIATION FOR THE MONTH" And V8_Norm(raw)<>"LIQUIDATION OF PCF FOR THE MONTH" And V8_Norm(raw)<>"CLOSING OF INPUT VAT FOR Q1 2026" Then
            a=V8_GJMap(raw)
            ok=True:v=V8_Number(ws.Cells(r,6).Value2,ok):If Not ok Then inv=inv+1
            If ok And Abs(v)>TOLERANCE Then
                If Len(a)=0 Then um=um+Abs(v):detail=detail&"GJ!"&ws.Cells(r,5).Address(False,False)&" -> "&raw&" debit"&vbCrLf Else Put d,a,cur,v,0#:gd=gd+v
            End If
            ok=True:v=V8_Number(ws.Cells(r,7).Value2,ok):If Not ok Then inv=inv+1
            If ok And Abs(v)>TOLERANCE Then
                If Len(a)=0 Then um=um+Abs(v):detail=detail&"GJ!"&ws.Cells(r,5).Address(False,False)&" -> "&raw&" credit"&vbCrLf Else Put d,a,cur,0#,v:gc=gc+v
            End If
        End If
    Next r
    Set Matrix=d
End Function

Public Function CalculateMonthlyNet(ByVal accountName As String,ByVal monthNumber As Long,ByVal yearNumber As Long) As Double
    Dim d As Object,gd As Double,gc As Double,um As Double,inv As Long,s As String,a As Variant,k As String
    Set d=Matrix(yearNumber,gd,gc,um,inv,s):k=V8_Norm(accountName)&"|"&monthNumber
    If d.Exists(k) Then a=d(k):CalculateMonthlyNet=CDbl(a(0))-CDbl(a(1))
End Function

Public Function CalculateEndingBalanceFull(ByVal accountName As String,ByVal yearNumber As Long,Optional ByVal openingBalance As Double=0) As Double
    Dim m As Long:CalculateEndingBalanceFull=openingBalance
    For m=1 To 12:CalculateEndingBalanceFull=CalculateEndingBalanceFull+CalculateMonthlyNet(accountName,m,yearNumber):Next m
End Function

Public Function ValidateGLConsistency(Optional ByVal yearNumber As Long=2026) As Boolean
    Dim d As Object,gd As Double,gc As Double,um As Double,inv As Long,detail As String,ws As Worksheet,r As Long,ok As Boolean
    Set d=Matrix(yearNumber,gd,gc,um,inv,detail)
    On Error Resume Next:Set ws=ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then Set ws=ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)):ws.Name=AUDIT_SHEET
    On Error GoTo 0:r=ws.Cells(ws.Rows.Count,1).End(xlUp).Row+1
    ok=(Abs(gd-gc)<=TOLERANCE And um<=TOLERANCE And inv=0)
    ws.Cells(r,1).Value=Now:ws.Cells(r,2).Value=IIf(ok,"PASS","HOLD")
    ws.Cells(r,3).Value="v8.0 | CDJ+CRJ+GJ | Debit="&Format$(gd,"0.00")&" | Credit="&Format$(gc,"0.00")&" | Difference="&Format$(gd-gc,"0.00")&" | Unmapped="&Format$(um,"0.00")&" | Invalid="&inv
    ws.Cells(r,4).Value=detail
    ValidateGLConsistency=ok
End Function

Public Function BuildV8CalcSheet(Optional ByVal yearNumber As Long=2026) As Boolean
    Dim d As Object,gd As Double,gc As Double,um As Double,inv As Long,detail As String,ws As Worksheet,accounts As Object,a As Variant,m As Long,r As Long,k As String,x As Variant,bal As Double
    Set d=Matrix(yearNumber,gd,gc,um,inv,detail):Set accounts=V8_GLAccounts()
    On Error Resume Next:Set ws=ThisWorkbook.Worksheets(CALC_SHEET)
    If ws Is Nothing Then Set ws=ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)):ws.Name=CALC_SHEET
    On Error GoTo 0:ws.Cells.Clear:ws.Range("A1:E1").Value=Array("Account Title","Month","Debit","Credit","Ending Balance"):r=2
    For Each a In accounts.Items
        bal=0
        For m=1 To 12
            k=V8_Norm(CStr(a))&"|"&m:If d.Exists(k) Then x=d(k) Else x=Array(0#,0#)
            bal=bal+CDbl(x(0))-CDbl(x(1))
            ws.Cells(r,1).Value=a:ws.Cells(r,2).Value=m:ws.Cells(r,3).Value=x(0):ws.Cells(r,4).Value=x(1):ws.Cells(r,5).Value=bal:r=r+1
        Next m
    Next a
    ws.Columns("A:E").AutoFit:BuildV8CalcSheet=(r-2=46*12)
End Function

Public Sub RefreshGL(Optional ByVal yearNumber As Long=2026)
    If Not ValidateGLConsistency(yearNumber) Then MsgBox "v8.0 HOLD: unmapped/invalid posting activity. Existing GL was not overwritten.",vbExclamation:Exit Sub
    If Not BuildV8CalcSheet(yearNumber) Then MsgBox "v8.0 HOLD: 46x12 matrix failed.",vbExclamation:Exit Sub
    MsgBox "v8.0 46x12 matrix built. Existing GL remains protected pending zero-unmapped certification.",vbInformation
End Sub

Public Sub RefreshAllGL(Optional ByVal yearNumber As Long=2026):RefreshGL yearNumber:End Sub
