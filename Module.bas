Attribute VB_Name = "Module"
Sub getInfo_Click()

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

'Macros developed to edit graph
'this is called from "Plot" button and private sub "CreateGraph"

Sub editGraph()
Dim wks As Worksheet
Dim cht As Chart

Set wks = ActiveWorkbook.Sheets("DataRaw")
wks.ChartObjects("Historical_Data").Select  'selecting graph from active sheet

Set cht = ActiveWorkbook.Sheets("DataRaw").ChartObjects("Historical_Data").Chart 'setting selected grapg to "cht" variable

'  WITH SELECTION EVERY CHANGE IN FORMAT IS APPLIED AS FOLLOWS

cht.HasTitle = True 'graph title
cht.ChartTitle.Text = "HISTORICAL DATA"
With cht.ChartTitle.Font
    .Size = 18
    .Bold = True
    .Color = RGB(68, 114, 196)
End With

'chart gridlines
cht.Axes(xlCategory).HasMajorGridlines = False
cht.Axes(xlCategory).HasMinorGridlines = False
cht.Axes(xlValue).HasMajorGridlines = False
cht.Axes(xlValue).HasMinorGridlines = False


'X axis name
cht.Axes(xlCategory, xlPrimary).HasTitle = True
cht.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Date"

'y-axis name and number format
cht.Axes(xlValue, xlPrimary).HasTitle = True
cht.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Adj Close, USD"
cht.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0.00"

'graph legend position
cht.HasLegend = True
cht.Legend.Position = xlLegendPositionTop

'editing line smooth, line color and mark color
cht.PlotArea.Select
    cht.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .Transparency = 0
        .Solid
    End With
    
cht.FullSeriesCollection(1).Smooth = True


'placing graph to another sheet (incompleted)

cht.Location Where:=xlLocationAsNewSheet


End Sub

'it is called from YahooScrap private sub with "Import Data" button
Sub clearData() 'function created to clear prior data stored in sheets

Dim wbQuery As Worksheet

Set wbQuery = Workbooks(ThisWorkbook.Name).Worksheets("WebQuery")
lastRow = wbQuery.Range("A" & Rows.Count).End(xlUp).Row
wbQuery.Range("A4:F" & lastRow).ClearContents
End Sub

