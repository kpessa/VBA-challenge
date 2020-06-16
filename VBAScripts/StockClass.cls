Option Explicit

' Custom StockClass

' Properties
Public tickerSymbol As String
Public openingPrice As Double
Public closingPrice As Double
Public totalStockVolume As Double

' Methods
Public Function yearlyChange()
    yearlyChange = closingPrice - openingPrice
End Function

'@Ignore ProcedureNotUsed
Public Function percentChange()
    If openingPrice <> 0 Then percentChange = yearlyChange / openingPrice
End Function