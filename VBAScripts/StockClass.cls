VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Custom StockClass

' Properties
Public tickerSymbol As String
Public openingPrice As Double
Public closingPrice As Double
Public totalStockVolume As Double

' Methods
Public Function yearlyChange() As Double
    yearlyChange = closingPrice - openingPrice
End Function

'@Ignore ProcedureNotUsed
Public Function percentChange() As Double
    If openingPrice <> 0 Then percentChange = yearlyChange / openingPrice
End Function
