Public startTime As Double

Sub Main()
    
    Dim ws As Worksheet
    startTime = Timer
    
    Call EnablePerformance
    
    For Each ws In Workbooks("Multiple_year_stock_data.xlsx").Worksheets
        
        Debug.Print vbNewLine & "Time started year " & ws.Name & ": " & Timer - startTime & " s"
        
        ' Macro_NotUsingArray ws
        Macro ws
        Challenge ws
        
    Next
    
    Call ResetPerformance
    
    Debug.Print vbNewLine & "Time ended: " & Timer - startTime & " s"
    
End Sub

Sub Challenge(ws As Worksheet)

    ' Setting header row
    ws.Range("P1:Q1") = Array("Ticker", "Value")
    
    ' Setting row descriptions
    ws.Range("O2:O4") = Excel.WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    
    ' Using excel's built-in function "XLOOKUP" and "Max/Min" to find max, min values
    ' 1. Greatest % increase
    ws.Range("P2") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), ws.Range("I:I"))
    ws.Range("Q2") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), ws.Range("K:K"))
    ws.Range("Q2").NumberFormat = "0.00%" ' set number format to percentage
    ' 2. Greatest % decrease
    ws.Range("P3") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), ws.Range("I:I"))
    ws.Range("Q3") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), ws.Range("K:K"))
    ws.Range("Q3").NumberFormat = "0.00%" ' set number format to percentage
    ' 3. Greatest Total Volume
    ws.Range("P4") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), ws.Range("I:I"))
    ws.Range("Q4") = Excel.WorksheetFunction.XLookup(Excel.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), ws.Range("L:L"))
    ws.Range("Q4").NumberFormat = "0.00E+00" ' set number format to scientific
    
    Debug.Print "Finished Challenge for " & ws.Name & ": " & Timer - startTime & " s"
    
End Sub

Sub Macro(ws As Worksheet)
    Dim stockCollection As New Collection
    Dim stock As StockClass
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Value
    
    Debug.Print "Loaded " & UBound(arr, 1) & " rows into an array at: " & Timer - startTime & " s"
    
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
    ws.Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    ' Setting column K or "Percent Change" to Percent number format
    ws.Columns("K:K").NumberFormat = "0.00%"
    
    For i = 1 To stockCollection.Count
    
        ws.Range("I" & i + 1) = stockCollection(i).tickerSymbol
        ws.Range("J" & i + 1) = stockCollection(i).yearlyChange
        ws.Range("K" & i + 1) = stockCollection(i).percentChange
        ws.Range("L" & i + 1) = stockCollection(i).totalStockVolume
        
        If stockCollection(i).yearlyChange > 0 Then
            ws.Range("J" & i + 1).Interior.Color = vbGreen
        ElseIf stockCollection(i).yearlyChange < 0 Then
            ws.Range("J" & i + 1).Interior.Color = vbRed
        End If
        
    Next
    
    Debug.Print "Outputted " & stockCollection.Count & " stocks to summary table at: " & Timer - startTime & " s"
    
End Sub

Sub Macro_NotUsingArray(ws as Worksheet)
    
    Dim stockCollection As New Collection
    Dim stock As StockClass
    
    For i = 2 To ws.Range("A1").End(xlDown).Row
            ' 1.) New Stock: Use lookbehind to see if it's a new stock
            If ws.Range("A" & i) <> ws.Range("A" & i - 1) Then
                Set stock = New StockClass
                stock.tickerSymbol = ws.Range("A" & i)
                stock.openingPrice = ws.Range("C" & i)
                stock.totalStockVolume = ws.Range("G" & i)
            ' 2.) Last Record of Stock: Use lookahead to see if it's the last record of the stock
            ElseIf ws.Range("A" & i) <> ws.Range("A" & i + 1) Then
                stock.closingPrice = ws.Range("F" & i)
                stock.totalStockVolume = stock.totalStockVolume + CDbl(ws.Range("G" & i))
                stockCollection.Add stock, stock.tickerSymbol
            ' 3.) In the middle: If in the middle, continue aggregating totalStockVolume
            Else
                stock.totalStockVolume = stock.totalStockVolume + CDbl(ws.Range("G" & i))
            End If
    Next
    
    Debug.Print "Finished processing " & ws.Range("A1").End(xlDown).Row & " rows at: " & Timer - startTime & " s"
    
    For i = 1 To stockCollection.Count
        ws.Range("I" & i + 1) = stockCollection(i).tickerSymbol
        ws.Range("J" & i + 1) = stockCollection(i).yearlyChange
        ws.Range("K" & i + 1) = stockCollection(i).percentChange
        ws.Range("L" & i + 1) = stockCollection(i).totalStockVolume
    Next
    
    Debug.Print "Outputted " & stockCollection.Count & " stocks to summary table at: " & Timer - startTime & " s"
    
End Sub

Sub EnablePerformance()
    ' Performance
    Application.EnableAnimations = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Sub ResetPerformance()
    ' Reset Performance
    Application.EnableAnimations = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
