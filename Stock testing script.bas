Attribute VB_Name = "Module1"
Sub homework():

'create a loop running through  every row in column A looking for a change in ticker symbol.
'to find the yearly change in dollar amount and percent amount. How would i do this? when the ticker symbol changes,
'take the value to start, then when the ticker changes again, take the value of the cell above and do the calculations.
'Make a running total for all of the diferent ticker symbols.
'create a summary table that displays, ticker symbol, year change in price, percent change in price, total volume
'My variables will be what? Ticker symbols, running total for volume, begginging of the year value, end of the year value, end of row
'summary table variables are yearly change in price, percent change in price, total volume, and ticker symbol

'there are 261 rows for each ticker symbol, i can use that to determine the starting value by taking 260 away the ticker symbol changes... is there any other way to do this?
    'this does not work because in the yearly version, all tickers dont have same number of. Count the number of times we run loop since change in ticker
'to get the ending value, when the ticker changes, it will just be the value in the F column

For Each ws In Worksheets
    Dim Ticker As String
    Dim TotalVol As Variant
    Dim Bvalue As Double
    Dim Evalue As Double
    Dim LastRow As LongLong
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    Dim Ychange As Double
    Dim Pchange As Double
    Dim DaysSinceOpen As Integer
    DaysSinceOpen = 0



    For Row = 2 To LastRow
        'print the ticker symbols in summary table
        
        
        
        
        
        If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) And ws.Range("C" & Row).Value > 0 Then
        
            'set the ticker symbol
            Ticker = ws.Range("A" & Row).Value
            'print in table
            ws.Range("K" & SummaryTableRow).Value = Ticker
            'add last volume
            TotalVol = TotalVol + ws.Range("G" & Row).Value
            'add to summary table
            ws.Range("N" & SummaryTableRow).Value = TotalVol
            'Get first open value, Bvalue
            Bvalue = ws.Range("C" & Row - DaysSinceOpen).Value
            'Get last closing value, Evalue
            Evalue = ws.Range("F" & Row).Value
            'calculate yearly change
            Ychange = Evalue - Bvalue
            'Put Yearly change in summary table
            ws.Range("L" & SummaryTableRow).Value = Ychange
            
            
            'Calculate Percent Change
            Pchange = (Evalue - Bvalue) / Bvalue
            'place percent change in summary table
            ws.Range("M" & SummaryTableRow).Value = Pchange
            'change formatting to percentage
            ws.Range("M" & SummaryTableRow).NumberFormat = "0.00%"
            'Add conditionals for red and green if positive or negative
            'add row to table
            SummaryTableRow = SummaryTableRow + 1
            'reset the total volume
            TotalVol = 0
            DaysSinceOpen = 0
        ElseIf (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) And ws.Range("C" & Row - DaysSinceOpen).Value = 0 Then
            
            ws.Range("N" & SummaryTableRow).Value = "XXX"
            SummaryTableRow = SummaryTableRow + 1
            DaysSinceOpen = 0
            TotalVol = 0
        'if in same ticker we go from a 0 starting value to an end value
        ElseIf ws.Range("A" & Row).Value = ws.Range("A" & Row + 1).Value And ws.Range("C" & Row).Value = 0 Then
            DaysSinceOpen = 0
            TotalVol = 0
        Else
            TotalVol = TotalVol + ws.Range("G" & Row).Value
            DaysSinceOpen = DaysSinceOpen + 1
            
        
        
        End If
     
        
    
    Next Row

    For Row1 = 2 To LastRow

        If ws.Range("L" & Row1).Value > 0 Then
                ws.Range("L" & Row1).Interior.ColorIndex = 10
            ElseIf ws.Range("L" & Row1).Value < 0 Then
                ws.Range("L" & Row1).Interior.ColorIndex = 3
            End If

    Next Row1

Next ws

End Sub
