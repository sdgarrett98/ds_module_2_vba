Sub stocks()
    Dim ws As Worksheet ' used Xpert
    Dim i As Long
    Dim cell_vol As LongLong
    Dim vol_total As LongLong
    Dim ticker As String
    Dim k As Long
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim pct_change As Double
    
    Dim lastRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets ' used Xpert
        Dim greatest_inc As Double
        Dim greatest_dec As Double
        Dim greatest_vol As Double
        
        ' Set headers for leaderboards
        ' Set leaderboard column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Set greatest column headers
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        
        ' Set greatest row headers
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        ' Initialize greatest values for each worksheet
        greatest_inc = 0
        greatest_dec = 0
        greatest_vol = 0
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        vol_total = 0
        k = 2
        
        ticker_open = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            
            ' Calculate price change and percentage change
            ' ticker_close = ws.Cells(i, 6).Value
            ' price_change = ticker_close - ticker_open
            ' pct_change = price_change / ticker_open
            
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                vol_total = vol_total + cell_vol
                
                ticker_close = ws.Cells(i, 6).Value
                price_change = ticker_close - ticker_open
                If (ticker_open = 0) Then
                    pct_change = 0
                Else
                    pct_change = price_change / ticker_open
                End If
    
                ws.Cells(k, 9).Value = ticker
                
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 10).NumberFormat = "0.00"
                
                ws.Cells(k, 11).Value = pct_change
                ws.Cells(k, 11).NumberFormat = "0.00%"
                
                ws.Cells(k, 12).Value = vol_total
                
                If price_change > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
                ElseIf price_change < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3
                End If
                
                ' reset
                vol_total = 0
                k = k + 1
                ticker_open = ws.Cells(i + 1, 3).Value
            Else
                vol_total = vol_total + cell_vol
            End If
            
            ' Check for greatest increase
            If pct_change > greatest_inc Then
                greatest_inc = pct_change
                ws.Cells(2, 17).Value = ticker
                ws.Cells(2, 18).Value = greatest_inc
                ws.Cells(2, 18).NumberFormat = "0.00%"
            End If
            
            ' Check for greatest decrease
            If pct_change < greatest_dec Then
                greatest_dec = pct_change
                ws.Cells(3, 17).Value = ticker
                ws.Cells(3, 18).Value = greatest_dec
                ws.Cells(3, 18).NumberFormat = "0.00%"
            End If
            
            ' Check for greatest total volume
            If vol_total > greatest_vol Then
                greatest_vol = vol_total
                ws.Cells(4, 17).Value = ticker
                ws.Cells(4, 18).Value = greatest_vol
            End If
            
        Next i
    Next ws
End Sub

