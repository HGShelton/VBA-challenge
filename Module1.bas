Attribute VB_Name = "Module1"
Sub Stocks()
    
    'Set variables
    Dim ticker As String
    Dim ws As Worksheet
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim summary_table As Double
    
    'loop through worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    'Set columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Set summary table
    summary_table = 2
    
    'Loop through all tickers
    For i = 2 To ws.UsedRange.Rows.Count
    
    If i = 2 Then
        open_price = ws.Cells(i, 3).Value
    
    End If
    
        'Check if ticker has changed, calculate yearly change
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
            'Set values
            ticker = ws.Cells(i, 1)
            close_price = ws.Cells(i, 6).Value
            yearly_change = close_price - open_price
            percent_change = (close_price - open_price) / open_price
            total_volume = total_volume + ws.Cells(i, 7).Value
            
                'Print values
                ws.Cells(summary_table, 9) = ticker
                ws.Cells(summary_table, 10).Value = yearly_change
                ws.Cells(summary_table, 11).Value = percent_change
                ws.Cells(summary_table, 12).Value = total_volume
                summary_table = summary_table + 1
            
            'reset values
            open_price = ws.Cells(i + 1, 3).Value
            year_close = 0
            total_volume = 0
            
        Else
        total_volume = total_volume + ws.Cells(i, 7).Value
                
        End If
     
    Next i
  
    'format percentage
    ws.Columns("K").NumberFormat = "0.00%"
  
    'format colors
    For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(j, 10).Interior.ColorIndex = 4
  
        End If
  
    Next j
  
    'Greatest Values
    Dim ticker_inc, ticker_dec, ticker_volume As String
    Dim percent_inc, percent_dec, great_volume As Double
    
    'set columns
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'set value
    percent_inc = 0
    percent_dec = 0
    great_volume = 0

    'loop
    For k = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(k, 11).Value > percent_inc Then
            percent_inc = ws.Cells(k, 11).Value
                ticker_inc = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 11).Value < percent_dec Then
            percent_dec = ws.Cells(k, 11).Value
                ticker_dec = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 12).Value > great_volume Then
            great_volume = ws.Cells(k, 12).Value
                ticker_volume = ws.Cells(k, 9).Value
        End If
        
    'print tickers
     ws.Cells(2, 16) = ticker_inc
     ws.Cells(3, 16) = ticker_dec
     ws.Cells(4, 16) = ticker_volume
     
     'set values
     ws.Cells(2, 17) = percent_inc
     ws.Cells(3, 17) = percent_dec
     ws.Cells(4, 17) = great_volume
                
    Next k
    
    'format values
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
  Next ws
End Sub

