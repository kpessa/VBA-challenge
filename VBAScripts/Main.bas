Option Explicit

Public Sub Main()
    
    Dim startTime As Double
    startTime = Timer
    
    EnablePerformance
    
    Workbooks.Open ThisWorkbook.Path & "\VBAStocks\Multiple_year_stock_data.xlsx"
    Dim ws As Worksheet
    For Each ws In Workbooks("Multiple_year_stock_data.xlsx").Worksheets
        
        Debug.Print vbNewLine & "Time started year " & ws.Name & ": " & Timer - startTime & " s"
        
        'Call Macro_NotUsingArray startTime
        Macro ws, startTime
        Challenge ws, startTime
        
    Next
    
    ResetPerformance
    
    Debug.Print vbNewLine & "Time ended: " & Timer - startTime & " s"
    
End Sub

Private Sub Challenge(ws As Worksheet, ByVal startTime As Double)

    ' Setting header row
    ws.Range("P1:Q1").Value = Array("Ticker", "Value")
    
    ' Setting row descriptions
    ws.Range("O2:O4").Value = Excel.WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    
    ' Using excel's built-in function "XLOOKUP" and "Max/Min" to find max, min values
    ' 1. Greatest % increase
    ws.Range("P2").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), ws.Range("I:I"))
    ws.Range("Q2").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), ws.Range("K:K"))
    ws.Range("Q2").NumberFormat = "0.00%" ' set number format to percentage
    ' 2. Greatest % decrease
    ws.Range("P3").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), ws.Range("I:I"))
    ws.Range("Q3").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), ws.Range("K:K"))
    ws.Range("Q3").NumberFormat = "0.00%" ' set number format to percentage
    ' 3. Greatest Total Volume
    ws.Range("P4").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), ws.Range("I:I"))
    ws.Range("Q4").Value = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), ws.Range("L:L"))
    ws.Range("Q4").NumberFormat = "0.00E+00" ' set number format to scientific
    
    Debug.Print "Finished Challenge for " & ws.Name & ": " & Timer - startTime & " s"
    
End Sub

Private Sub Macro(ws As Worksheet, ByVal startTime As Double)
    Dim stockCollection As Collection: Set stockCollection = New Collection
    Dim stock As StockClass
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Value
    
    Debug.Print "Loaded " & UBound(arr, 1) & " rows into an array at: " & Timer - startTime & " s"
    
    Dim i As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
            ' 1.) New Stock: Use lookbehind to see if it's a new stock
            If arr(i, 1) <> arr(i - 1, 1) Then
                Set stock = New StockClass
                stock.tickerSymbol = arr(i, 1)
                stock.openingPrice = arr(i, 3)
                stock.totalStockVolume = arr(i, 7)
            ' 2.) Last Stock
            ElseIf i = UBound(arr, 1) Then
                stock.closingPrice = arr(i, 6)
                stock.totalStockVolume = stock.totalStockVolume + CDbl(arr(i, 7))
                stockCollection.Add stock, stock.tickerSymbol
            ' 3.) Last Record of Stock: Use lookahead to see if it's the last record of the stock
            ElseIf arr(i, 1) <> arr(i + 1, 1) Then
                stock.closingPrice = arr(i, 6)
                stock.totalStockVolume = stock.totalStockVolume + CDbl(arr(i, 7))
                stockCollection.Add stock, stock.tickerSymbol
            ' 4.) In the middle: If in the middle, continue aggregating totalStockVolume
            Else
                stock.totalStockVolume = stock.totalStockVolume + CDbl(arr(i, 7))
            End If
    Next
    
    ' Setting Header Row
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    ' Setting column K or "Percent Change" to Percent number format
    ws.Columns("K:K").NumberFormat = "0.00%"
    
    For i = 1 To stockCollection.Count
    
        ws.Range("I" & i + 1).Value = stockCollection(i).tickerSymbol
        ws.Range("J" & i + 1).Value = stockCollection(i).yearlyChange
        ws.Range("K" & i + 1).Value = stockCollection(i).percentChange
        ws.Range("L" & i + 1).Value = stockCollection(i).totalStockVolume
        
        If stockCollection(i).yearlyChange > 0 Then
            ws.Range("J" & i + 1).Interior.Color = vbGreen
        ElseIf stockCollection(i).yearlyChange < 0 Then
            ws.Range("J" & i + 1).Interior.Color = vbRed
        End If
        
    Next
    
    Debug.Print "Outputted " & stockCollection.Count & " stocks to summary table at: " & Timer - startTime & " s"
    
End Sub

' Ended up not using
'Sub Macro_NotUsingArray(startTime As Double)
'
'    Dim stockCollection As New Collection
'    Dim stock As StockClass
'
'    Dim ws As Worksheet
'    Set ws = Workbooks("Multiple_year_stock_data.xlsx").Worksheets(1)
'
'    Dim i As Long
'    For i = 2 To ws.Range("A1").End(xlDown).Row
'            ' 1.) New Stock: Use lookbehind to see if it's a new stock
'            If ws.Range("A" & i).Value <> ws.Range("A" & i - 1).Value Then
'                Set stock = New StockClass
'                stock.tickerSymbol = ws.Range("A" & i).Value
'                stock.openingPrice = ws.Range("C" & i).Value
'                stock.totalStockVolume = ws.Range("G" & i).Value
'            ' 2.) Last Record of Stock: Use lookahead to see if it's the last record of the stock
'            ElseIf ws.Range("A" & i).Value <> ws.Range("A" & i + 1).Value Then
'                stock.closingPrice = ws.Range("F" & i).Value
'                stock.totalStockVolume = stock.totalStockVolume + CDbl(ws.Range("G" & i).Value)
'                stockCollection.Add stock, stock.tickerSymbol
'            ' 3.) In the middle: If in the middle, continue aggregating totalStockVolume
'            Else
'                stock.totalStockVolume = stock.totalStockVolume + CDbl(ws.Range("G" & i).Value)
'            End If
'    Next
'
'    Debug.Print "Finished processing " & ws.Range("A1").End(xlDown).Row & " rows at: " & Timer - startTime & " s"
'
'    For i = 1 To stockCollection.Count
'        ws.Range("I" & i + 1).Value = stockCollection(i).tickerSymbol
'        ws.Range("J" & i + 1).Value = stockCollection(i).yearlyChange
'        ws.Range("K" & i + 1).Value = stockCollection(i).percentChange
'        ws.Range("L" & i + 1).Value = stockCollection(i).totalStockVolume
'    Next
'
'    Debug.Print "Outputted " & stockCollection.Count & " stocks to summary table at: " & Timer - startTime & " s"
'
'End Sub

' Not used
'Sub ForLoop()
'    Dim GPI As Double: GPD As Double: GTV As Double
'    Dim GPI_ticker As String: GPD_ticker As String: GTV_ticker As String
'
'    Dim ws As Worksheet
'    Set ws = Workbooks("Multiple_year_stock_data.xlsx").Worksheets(1)
'
'    Dim i As Long
'    For i = 2 To ws.Range("I2").End(xlDown).Row
'        ' 1. Conditional for GPI
'        If ws.Range("K" & i) > GPI Then
'            GPI = ws.Range("K" & i)
'            GPI_ticker = ws.Range("I" & i)
'        End If
'
'        ' 2. Conditional for greatestPercDecrease
'        If ws.Range("K" & i) < GPD Then
'            GPD = ws.Range("K" & i)
'            GPD_ticker = ws.Range("I" & i)
'        End If
'
'        ' 3. Conditional for greatestTotalVolume
'        If ws.Range("L" & i) > GTV Then
'            GTV = ws.Range("L" & i)
'            GTV_ticker = ws.Range("I" & i)
'        End If
'
'    Next
'
'End Sub

Private Sub EnablePerformance()
    ' Performance
    Application.EnableAnimations = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub ResetPerformance()
    ' Reset Performance
    Application.EnableAnimations = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub