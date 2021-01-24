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
    
    lastRow = vbSheetDest.Range("F" & Rows.Count).End(xlUp).Row - 2 'get the last row of origin data
    vbSheetDest.Range("H2:P2").ColumnWidth = 17
    
        
        For i = 4 To lastRow
        
            vbSheetDest.Range("H" & i) = vbSheetDest.Range("F" & i).Value / vbSheetDest.Range("F" & i - 1).Value
        
        Next i
        
        '**** LOG OF RISK (ri) CALCULATION ******
        
        For i = 4 To lastRow
            
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
        
        For i = 4 To lastRow
        
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

lastRow = thisWorksheet.Range("A" & Rows.Count).End(xlUp).Row - 2 'get the last row of origin data


thisWorksheet.Range("A2:P" & lastRow).ClearContents

thisWorksheet.Range("R5:R8").ClearContents



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

