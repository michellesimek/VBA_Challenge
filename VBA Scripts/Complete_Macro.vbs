Sub stocks():
'Define variables
Dim lastrow As Long
Dim lastcol As Long
Dim ticker As String
Dim ws As Worksheet
Dim total As LongLong
Dim summaryrow As Double

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
    ws.Cells(2, lastcol + 7) = "Greatest % Increase"
    ws.Cells(3, lastcol + 7) = "Greatest % Decrease"
    ws.Cells(4, lastcol + 7) = "Greatest Total Volume"
    ws.Cells(1, lastcol + 8) = "Ticker"
    ws.Cells(1, lastcol + 9) = "Value"
    
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
            'add total to Total Stock Volume column
            ws.Cells(summaryrow, 11).Value = total
                    
            'reset values
            summaryrow = summaryrow + 1
            total = 0
            inital_open = ws.Cells(Row + 1, 3).Value
                
        End If
    
    Next Row
    
    'determine max total stock volume
    yearly_change = ws.Range("K:K")
    max_val = WorksheetFunction.max(yearly_change)
    ws.Cells(4, 16) = max_val
                
    'determine greatest % increase and decrease
    percent_change = ws.Range("J:J")
    greatest_incr = WorksheetFunction.max(percent_change)
    greatest_decr = WorksheetFunction.min(percent_change)
    ws.Cells(2, 16) = greatest_incr
    ws.Cells(3, 16) = greatest_decr
    
    'loop through rows
    For Row = 2 To lastrow
    
        'if statement to find max total stock volume and return ticker symbol
        If ws.Cells(Row, 11).Value = max_val Then
            ws.Cells(4, 15).Value = ws.Cells(Row, 8).Value

            'Exit For loop once found   
            Exit For
        
        End If
    
    Next Row

    'loop through rows   
    For Row = 2 To lastrow
        
        'If statement to find greatest % increase value and return ticker symbol
        If ws.Cells(Row, 10).Value = greatest_incr Then
            ws.Cells(2, 15).Value = ws.Cells(Row, 8).Value
        
            'Exit For loop once found
            Exit For
        
        End If
    
    Next Row

    'loop through rows  
    For Row = 2 To lastrow

        'If statement to find greatest % decrease value and return ticker symbol
        If ws.Cells(Row, 10).Value = greatest_decr Then
            ws.Cells(3, 15).Value = ws.Cells(Row, 8).Value
        
            'Exit For loop once found
            Exit For
        
        End If
              
     Next Row
                
Next ws

End Sub