VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'The model estimates the variation over time of financial instruments.
'It assumes these instruments (such as stocks or futures) will have a lognormal distribution of prices.
'Using this assumption and factoring in other important variables, the equation derives the price of a call option
Private Sub B_Sholes_Click()
Dim vbSheetDest As Worksheet
Dim ri As Double
Dim n As Integer
Dim d1, d2 As Double
Dim C0, P0 As Double

Set vbSheetDest = Worksheets("DataRaw") 'sheet name of destination workbook

'*** CALLS AND PUTS CALCULATIONS ****

S = vbSheetDest.Range("R2").Value
X = vbSheetDest.Range("R3").Value
T = vbSheetDest.Range("R4").Value
AV = vbSheetDest.Range("O3").Value
r = vbSheetDest.Range("P3").Value

        If S = "" Then

            MsgBox "Ingrese Subyacente"
            vbSheetDest.Range("R2").Select

        ElseIf X = "" Then

            MsgBox "Ingrese Strike"
            vbSheetDest.Range("R3").Select

        ElseIf T = "" Then

            MsgBox "Ingrese Tiempo de Vencimiento"
            vbSheetDest.Range("R4").Select
            
        ElseIf r = "" Then
    
            MsgBox "Ingrese Tasa libre de riesgo"
            vbSheetDest.Range("P3").Select
            
        ElseIf AV = "" Then
            
            MsgBox "Calcule Volatilitad Anual"
            vbSheetDest.Range("O3").Select
    
        Else

        d1 = (Log(S / X) + (r + AV ^ 2 / 2) * T) / (AV * Sqr(T))

        vbSheetDest.Range("Q5").Value = "d1"
        vbSheetDest.Range("R5").Value = d1


        d2 = d1 - AV * Sqr(T)

        vbSheetDest.Range("Q6").Value = "d2"
        vbSheetDest.Range("R6").Value = d2

        C0 = S * Application.NormSDist(d1) - X * Exp(-r * T) * Application.NormSDist(d2)

        vbSheetDest.Range("Q7").Value = "Call"
        vbSheetDest.Range("R7").Value = C0

        P0 = X * Exp(-r * T) * Application.NormSDist(-d2) - S * Application.NormSDist(-d1)
        vbSheetDest.Range("Q8").Value = "Put"
        vbSheetDest.Range("R8").Value = P0
    End If
    
End Sub

Private Sub calibrate_Click()
Dim vbSheetDest As Worksheet
Dim ri As Double
Dim n As Integer
Dim d1, d2 As Double
Dim C0, P0 As Double


'**** ROI FROM CLOSED PRICES CALCULATION *****
Set vbSheetDest = Worksheets("DataRaw") 'sheet name of destination workbook

  lastRow = vbSheetDest.Cells(vbSheetDest.Rows.Count, "A").End(xlUp).Row
    
    'lastRow = vbSheetDest.Range("A" & Rows.Count).End(xlUp).Row 'get the last row of origin data
    vbSheetDest.Range("H2:P2").ColumnWidth = 17
    vbSheetDest.Range("H5") = 0
        
        For i = 6 To lastRow
        
        vbSheetDest.Range("H" & i) = vbSheetDest.Range("F" & i).Value / vbSheetDest.Range("F" & i - 1).Value
        
        Next i
        
        '**** LOG OF RISK (ri) CALCULATION ******
        
        For i = 6 To lastRow
            
            vbSheetDest.Range("I" & i).Value = Log(vbSheetDest.Range("H" & i).Value)
        
        Next i
        
        vbSheetDest.Range("H2").Value = "Daily Reveneu (ri)" 'ROI title
        vbSheetDest.Range("I2").Value = "Log ri" ' Log(ri) title
        vbSheetDest.Range("K2").Value = "Rho" 'Rho Title (Deriva)
        
        vbSheetDest.Range("J2").Value = "Nº Returns" 'Nº retorns title
        n = lastRow 'Nº of Returns
        vbSheetDest.Range("J3").Value = n
        
        
        '**** DERIVA (rho) CALCULATION *****
        
        rho = (Application.WorksheetFunction.Sum(vbSheetDest.Range("I3:I" & lastRow))) * (1 / n)
        vbSheetDest.Range("K3").Value = rho
        
        
        vbSheetDest.Range("L2").Value = "Ln(ri-rho)"
        vbSheetDest.Range("M2").Value = "Ln(ri-rho)^2"
        vbSheetDest.Range("N2").Value = "Sigma"
        
        For i = 6 To lastRow
        
            vbSheetDest.Range("L" & i).Value = Log(vbSheetDest.Range("H" & i).Value - rho)
            vbSheetDest.Range("M" & i).Value = (vbSheetDest.Range("L" & i).Value) ^ 2
        
        Next i
        
        ''**** VOLATILITY (sigma) CALCULATION ****
        
        sigma = Sqr((Application.WorksheetFunction.Sum(vbSheetDest.Range("M3:M" & lastRow))) * (1 / n))
        vbSheetDest.Range("N3").Value = sigma
        
        ''**** ANNUALIZED VOLATLILITY CALCULATION BY ASSUMING BROWNIAN BEHAVIOUR ****
        
        ''****Tagliafichi, R. A. (2008). La volatilidad, su cálculo y su incidencia en los precios de los derivados. Palermo Business Review, (2). Pg. 7.****
        
        'S = S1*365^(1/2)
        'S1 = sigma
        'S = AV (annualized volatility)
        '
        vbSheetDest.Range("O2").Value = "Annual Sigma"
        AV = sigma * 365 ^ (0.5)
        vbSheetDest.Range("O3").Value = AV
        
        vbSheetDest.Range("P2").Value = "Free Annual Risk" 'It should be given by Central Bank
        
        vbSheetDest.Range("Q2").Value = "Stock Price (today)" 'Look for price in Yahoo Finance o any other financial page
        vbSheetDest.Range("Q3").Value = "Strike Price (future)" 'set feasible future price by user
        vbSheetDest.Range("Q4").Value = "Time to maturity"
        'vbSheetDest.Range("R3").Value = "Normal Distribution"

End Sub

Private Sub Clear_Click()
Dim thisWorksheet As Worksheet

Set thisWorksheet = Workbooks(ThisWorkbook.Name).Worksheets("DataRaw")
lastRow = thisWorksheet.Range("A" & Rows.Count).End(xlUp).Row 'get the last row of origin data
thisWorksheet.Range("A5:P" & lastRow).ClearContents
thisWorksheet.Range("R5:R8").ClearContents
thisWorksheet.Range("B2").ClearContents
thisWorksheet.Range("D2").ClearContents
thisWorksheet.Range("F2").ClearContents
thisWorksheet.Range("J3").ClearContents
thisWorksheet.Range("K3").ClearContents
thisWorksheet.Range("N3").ClearContents
thisWorksheet.Range("O3").ClearContents


End Sub


Private Sub getInfo_Click()

Dim vbBookOrigen As Workbook
Dim vbSheetOrigin As Worksheet
Dim route As String
Dim vbBookDest As Workbook
Dim vbSheetDest As Worksheet

'IMPORTANT: CHANGE route VARIABLE TO ABSOLUTE LOCATION WHERE ORIGIN XLSX FILE IS PLACED, EXAMPLE: c:\..\origin.xlsx
route = "C:\Users\javie\OneDrive\Documentos\proyectos\Black-Sholes\IBE.MC.xlsx" 'get exactly route of origin information

Set vbBookOrigin = Workbooks.Open(route)
Set vbSheetOrigin = vbBookOrigin.Worksheets("IBE.MC") 'sheet name of origin workbook
Set vbBookDest = Workbooks(ThisWorkbook.Name)
Set vbSheetDest = vbBookDest.Worksheets("DataRaw") 'sheet name of destination workbook


lastRow = vbSheetOrigin.Range("A" & Rows.Count).End(xlUp).Row - 2 'get the last row of origin data

vbSheetOrigin.Range("A1:F" & lastRow).Copy Destination:=vbSheetDest.Range("A2") 'copy data from origin workbook to destination sheet

Workbooks(vbBookOrigin.Name).Close Savechanges:=False 'closing origin data workbook
 
 vbSheetDest.Range("A2:F360").Select
    With Selection.Font
        .Name = "Segoe UI"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With


End Sub

Private Sub ScrapYahoo_Click()
Dim URL$, xml, table, Q&, Mat, iRow, i&, Ticker$
Dim StartDate As Date
Dim EndDate As Date
Dim mFactor#, r&
Dim lastRow

Ticker = [b2]: StartDate = [d2]: EndDate = [f2]

    If StartDate > EndDate Then
        MsgBox "Please, check dates": End
    End If

    Application.ScreenUpdating = False
    
URL = "https://query1.finance.yahoo.com/v7/finance/download/" & _
Ticker & "?period1=" & _
date_Yahoo(StartDate) & "&period2=" & _
date_Yahoo(EndDate) & "&interval=1d&events=history"

    Do
        Set xml = CreateObject("msXML2.XMLHttp")
        xml.Open "Put", URL, False
        xml.Send

        table = Split(xml.responseText, vbLf)
        Q = UBound(table) - 1
    Loop While Q = -1

ReDim Mat(1 To Q, 0 To 5)

For i = 1 To Q
  iRow = Split(table(i), ",")
  If iRow(6) <> "null" Then
    r = 1 + r
    mFactor = Val(Replace(iRow(5), ",", "")) / Val(Replace(iRow(4), ",", ""))
  
    Mat(r, 0) = DateValue(iRow(0))
    Mat(r, 1) = Round(Val(Replace(iRow(6), ",", "")), 5)
    Mat(r, 2) = Round(mFactor * Val(Replace(iRow(1), ",", "")), 4)
    Mat(r, 3) = Round(mFactor * Val(Replace(iRow(2), ",", "")), 4)
    Mat(r, 4) = Round(mFactor * Val(Replace(iRow(3), ",", "")), 4)
    Mat(r, 5) = Round(mFactor * Val(Replace(iRow(4), ",", "")), 4)
  End If
Next

ThisWorkbook.Worksheets("WebQuery").Range("A4:F" & Q) = Mat

Application.ScreenUpdating = True

Set vbSheetOrigin = ThisWorkbook.Worksheets("WebQuery") 'sheet name of origin workbook
Set vbSheetDest = ThisWorkbook.Worksheets("DataRaw") 'sheet name of destination workbook


vbSheetOrigin.Range("A4:F" & Q).Copy Destination:=vbSheetDest.Range("A5") 'copy data from origin workbook to destination sheet

Exit Sub

verErr:
Debug.Assert False
Resume

End Sub

Private Function date_Yahoo(iDate As Date) As Long
  date_Yahoo = 86400# * CLng(iDate) - 2209161600#
End Function

Private Sub testUserForm_Click()

ImportTickert.Show

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim Range_1, Range_2, Range_3 As Range
    
Set Range_1 = ThisWorkbook.Worksheets("DataRaw").Range("A5:A10")
Set Range_2 = ThisWorkbook.Worksheets("DataRaw").Range("H5:H10")
Set Range_3 = ThisWorkbook.Worksheets("DataRaw").Range("B2")

    ScrapYahoo.Enabled = False
    calibrate.Enabled = False
    B_Sholes.Enabled = False
    Clear.Enabled = False

        If Application.WorksheetFunction.CountA(Range_1) > 0 Then
            ScrapYahoo.Enabled = True
            calibrate.Enabled = True
            Clear.Enabled = True
        End If
        If Application.WorksheetFunction.CountA(Range_2) > 0 Then
            B_Sholes.Enabled = True
            Clear.Enabled = True
        End If
        
        If Application.WorksheetFunction.CountA(Range_3) > 0 Then
            ScrapYahoo.Enabled = True
        End If
End Sub
