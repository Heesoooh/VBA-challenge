Attribute VB_Name = "Module11"
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Sub tickerResult():

    'add formulas to run all the sheet
    
    Dim ws As Integer
        For ws = 1 To Worksheets.Count
        Worksheets(ws).Select
    

    'Set dimentions
    Dim totalVolume As Double
    Dim Row As Long
    Dim openPrice As Double
    Dim closePrice As Double
    
    
    'Name result column
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'variables to hold ticker name.
    tickerName = ""
    
    'variables to hold the total stock volume
    totalVolume = 0
    
    'variables to hold summary row
    SummaryRow = 2

    
    'variables to hold openPrice and closePrice
    openPrice = Cells(2, 3).Value
    closePrice = Cells(2, 6).Value
    
    'variable to start at zero for percentChange and yearlyChange
    Dim percentChange As Double, yearlyChange As Double
    
    'declare variables
    percentChange = 0
    
    yearlyChange = 0
    
    'use function to find the last row in the sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop from row in the column
    For Row = 2 To lastrow
    
        ' check to see if the ticker changes
        

        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                
                ' if the ticker changes, do...
        
                ' fist set the ticker name
                tickerName = Cells(Row, 1).Value
                
                ' add the last volume from the row
                totalVolume = totalVolume + Cells(Row, 7).Value
        
                ' add the ticker name to the I Column in the Summmry row
                Cells(SummaryRow, 9).Value = tickerName
                
                ' add the total Volume to the L Column in the Summary row
                Cells(SummaryRow, 12).Value = totalVolume
                
                
                ' reset the ticker name to 0
                totalVolume = 0
                
                ' calculate yearlyChange
                closePrice = Cells(Row, 6).Value
                
                ' add formula
                yearlyChange = closePrice - openPrice
                
                ' add values to summaryRow J column
                Cells(SummaryRow, 10).Value = yearlyChange
                
                ' Format the summaryRow J column to currency ($)
                Cells(SummaryRow, 10).NumberFormat = "$ 0.00"
                
                ' calculate percentChange
                percentChange = yearlyChange / openPrice
                
                ' add values to summaryRow K column
                Cells(SummaryRow, 11).Value = percentChange
                
                ' Format the summaryRow K column to percentage
                Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                ' Color code the yearlyChange column
                
                If Range("J" & SummaryRow).Value > 0 Then
                    Range("J" & SummaryRow).Interior.ColorIndex = 4
                    
                ElseIf Range("J" & SummaryRow).Value < 0 Then
                        Range("J" & SummaryRow).Interior.ColorIndex = 3
                
                Else
                    
                End If
    
                
                openPrice = Cells(Row + 1, 3).Value
                
                ' go to the next summary table row (add 1 on to the value of the summary row)
                SummaryRow = SummaryRow + 1
                         
                
        Else
        
        
        ' if the ticker stays the same, do....
        ' add on to the total volume from the G column
        totalVolume = totalVolume + Cells(Row, 7).Value
        
        End If
         
    Next Row
    
            ' Add Greatest % increase, Greatest % decrease and Greatest total volume"
        
                Dim maxPercent, minPercent, maxVolume As LongLong
                Range("N2").Value = "Greatest % increase"
                Range("N3").Value = "Greatest % decrease"
                Range("N4").Value = "Greatest Total Volume"
                Range("O1").Value = "Ticker"
                Range("P1").Value = "Value"
                
                lastRows = Cells(Rows.Count, 9).End(xlUp).Row
                
                maxPercent = WorksheetFunction.Max(Range("K2:K" & lastRows))
                maxTickerNameIndex = WorksheetFunction.Match(maxPercent, Range("K2:K" & lastRows), 0)
                Range("O2").Value = Range("I" & maxTickerNameIndex + 1).Value
                Range("P2").Value = maxPercent
                Cells(2, 16).NumberFormat = "0.00%"
                
                minPercent = WorksheetFunction.Min(Range("K2:K" & lastRows))
                minTickerNameIndex = WorksheetFunction.Match(minPercent, Range("K2:K" & lastRows), 0)
                Range("O3").Value = Range("I" & minTickerNameIndex + 1).Value
                Range("P3").Value = minPercent
                Cells(3, 16).NumberFormat = "0.00%"
                
                maxVolume = WorksheetFunction.Max(Range("L2:L" & lastRows))
                maxTickerNameIndex = WorksheetFunction.Match(maxVolume, Range("L2:L" & lastRows), 0)
                Range("O4").Value = Range("I" & maxTickerNameIndex + 1).Value
                Range("P4").Value = maxVolume
                Cells(4, 16).NumberFormat = "0.00 E+0"
    
    Next ws
    
    
End Sub
