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
              
            'reset values
            summaryrow = summaryrow + 1
            total = 0
                
        End If
    
    Next Row
                
Next ws

End Sub
