VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportTickert 
   Caption         =   "Input Ticker"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   OleObjectBlob   =   "ImportTickert.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ImportTickert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ExitButton_Click()
ImportTickert.Hide
End Sub

Private Sub SuccessButton_Click()
Dim vbSheetDest As Worksheet

Set vbSheetDest = Worksheets("DataRaw")

vbSheetDest.Range("B2").Value = TickerText
vbSheetDest.Range("D2").Value = StartDateText
vbSheetDest.Range("F2").Value = EndDateText

TickerText = ""
StartDateText = ""
EndDateText = ""


ImportTickert.Hide


End Sub

Private Sub TickerText_Change()
TickerText.Text = UCase(TickerText.Text)
TickerText.SelStart = Len(TickerText)
End Sub

