'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
'Your result should look as follows (note: all solution images are for 2015 data).
'1. For loop
'2. Conditional statement
'3. Looking into the next cell
'4. Formula to get the last row with data (i.e. rowCount = Cells(Rows.Count, “A”).End(xlUp).Row)
'create a for loop to start at row 2 and count up until rowCount.
'embed an if statement that compares the values in row after the current row (i.e. row i+1 <> row i)
'
'and an else statement that adds the values together in column 7

'Moderate
Sub VBA_of_WallStreet()

'Declare variables
Dim TickerSymbol As String
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim YearlyChange As Double
Dim PercentChange As String

Dim TotalStockVolume As Double
Dim RowCounter As String

'Iterate over all worksheets
For Each ws In Worksheets
ws.Activate

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
RowCounter = 2

Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Total Stock Volume"
Cells(1, 11).Value = "Yearly Change" 'K
Cells(1, 12).Value = "Percent Change" 'L

'Code below finds changes only for last day of year.
'Could not figure how to pull first opening price value from first opening date


    For i = 2 To lastrow
        'Row 1 is header row
        'Check to see if value in next cell has changed
        'If the value in the next cell is NOT EQUAL (ticker symbol has changed), read values and write to columns I and J. Reset value.
        'Else add values for stock symbol into TotalStockVolumn
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'Read
                TickerSymbol = Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                YearlyOpen = YearlyOpen + Cells(i, 3)
                YearlyClose = YearlyClose + Cells(i, 6)
                YearlyChange = YearlyOpen - YearlyClose
                PercentChange = (YearlyOpen / YearlyClose)
                'Write
                Range("I" & RowCounter).Value = TickerSymbol
                Range("J" & RowCounter).Value = TotalStockVolume
                Range("K" & RowCounter).Value = YearlyChange
                Range("L" & RowCounter).Value = PercentChange
                'Increase count
                RowCounter = RowCounter + 1
                TotalStockVolume = 0
            Else
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            End If
    Next i
    
    
Next ws

End Sub

