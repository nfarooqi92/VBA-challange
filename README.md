# VBA-challange
The VBA of Wall Street
Sub StockData()

'make loop for all sheets
For Each ws In Worksheets

    'initial variables and values
    Dim Ticker_Name As String
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Year_Change As Double
    Dim Year_Percentage As Double
    Dim Stock_Volume As LongLong
    Year_Open = ws.Range("C2").Value
    
    'Keep track of the location for each stock in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'inital table's last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'last row in the summary table
    LastRow_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'Make summary table titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'loop through all stocks
    For i = 2 To LastRow
        
        'Check if we are within the same stock, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set stock name
            Ticker_Name = ws.Cells(i, 1).Value
            
            'closing price
            Year_Close = ws.Cells(i, 6).Value
            
            'percentage change of the year
            Year_Change = Year_Close - Year_Open
            
            'Divide by 0 error fix
            If Year_Close = 0 Then
                Year_Percentage = 0
            Else
                Year_Percentage = Year_Change / Year_Close
            End If
            
            'add to the stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
            'fill in ticker symbol in the summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            'fill in yearly change in the summary table
            ws.Range("J" & Summary_Table_Row).Value = Year_Change
            
            'fill in percent change in the summary table
            ws.Range("K" & Summary_Table_Row).Value = Year_Percentage
            
            'fill in stock volume in the summary table
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
            
            'add a row to the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset to the opening price
            Year_Open = ws.Cells(i + 1, 3).Value
            
            'reset the stock volume
            Stock_Volume = 0
        
        'If the cell immediately following a row is the same brand...
        Else
            
            'Add to the Stock Volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i

    'color yearly change: positive green negative red
    For j = 2 To LastRow_Summary
        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
    Next j

    'BONUS
    'new table titles
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'loop through summary table
    For y = 2 To LastRow_Summary
    
        'for greatest percent increase
        If ws.Cells(y, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow_Summary)) Then
            ws.Range("P2").Value = ws.Cells(y, 9).Value
            ws.Range("Q2").Value = ws.Cells(y, 11).Value
        
        'for greatest percent decrease
        ElseIf ws.Cells(y, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow_Summary)) Then
            ws.Range("P3").Value = ws.Cells(y, 9).Value
            ws.Range("Q3").Value = ws.Cells(y, 11).Value
            
        'for greatest total volume
        ElseIf ws.Cells(y, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow_Summary)) Then
            ws.Range("P4").Value = ws.Cells(y, 9).Value
            ws.Range("Q4").Value = ws.Cells(y, 12).Value
        End If
        
    Next y

Next ws

End Sub
