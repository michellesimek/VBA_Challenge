'Loop through each worksheet in the workbook
For Each ws In Worksheets

    'Assign variables
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    total = 0
    summaryrow = 2
    inital_open = ws.Cells(2, 3).Value
    
    'Create new columns in each worksheets
    ws.Cells(1, lastcol + 1) = "Ticker"
    ws.Cells(1, lastcol + 2) = "Yearly Change"
    ws.Cells(1, lastcol + 3) = "Percent Yearly Change"
    ws.Cells(1, lastcol + 4) = "Total Stock Volume"
    
    'Loop through each row in each worksheet
    For Row = 2 To lastrow
        
        'Add up the total stock volume for all ticker symbols that are the same
        total = ws.Cells(Row, 7).Value + total
            'If statements for when ticker symbol in first column varies
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            'add ticker symbols to Ticker columm
            ws.Cells(summaryrow, 8).Value = ws.Cells(Row, 1).Value
            'add yearly change from opening price to closing price
            ws.Cells(summaryrow, 9).Value = ws.Cells(Row, 6).Value - inital_open
            If ws.Cells(summaryrow, 9).Value > 0 Then
                ws.Cells(summaryrow, 9).Interior.ColorIndex = 4
            Else
                ws.Cells(summaryrow, 9).Interior.ColorIndex = 3
            End If
            'add percent yearly change from opening price and closing price
            If inital_open = 0 Then
                ws.Cells(summaryrow, 10).Value = ws.Cells(Row, 6).Value
            Else
                ws.Cells(summaryrow, 10).Value = Round(((ws.Cells(summaryrow, 9) / inital_open) * 100), 2) & " %"
            End If
                    
            'reset values
            summaryrow = summaryrow + 1
            total = 0
            inital_open = ws.Cells(Row + 1, 3).Value
                
        End If
    
    Next Row
                
Next ws

End Sub
