Attribute VB_Name = "StockAnalyzerAllSheets"
Sub StockAnalyzer()

For Each ws In Worksheets 'Run on each sheet one at a time

'Set up cells for results to land in--------------------------------------------
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Make sure data is sorted by date and ticker name (ignores the headers)
Range("A:G").Sort Key1:=Range("B1"), Order1:=xlAscending
Range("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending

'Determine the last non-empty row to use as a count for loop iteration
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'----------------------------------------------------------------------------------------


'Cycle through the list of unique ticker names and accumulate/compute data for each unique ticker name----------------
'Set variables for results to be calculated
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim YearStartValue As Double
    Dim YearEndValue As Double
    Dim CurrentTicker As String
    Dim NextTicker As String
    Dim UniqueTickerCount As Double
    
    TotalStockVolume = 0
    UniqueTickerCount = 0
    YearStartValue = Cells(2, 3).Value 'Initializes this value to the open of the first ticker listed
    

For i = 2 To LastRow
    CurrentTicker = ws.Cells(i, 1).Value 'Selects the ticker name of the current row
    NextTicker = ws.Cells(i + 1, 1).Value 'Looks at the ticker name of the next row
    
    If CurrentTicker <> NextTicker Then 'This is TRUE when the name on the next row is different, meaning this is the last row of the current unique ticker name
        UniqueTickerCount = UniqueTickerCount + 1
        
        ws.Cells(UniqueTickerCount + 1, 9).Value = CurrentTicker 'Put this name into the summary table
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        ws.Cells(UniqueTickerCount + 1, 12).Value = TotalStockVolume 'Record the total stock volume in the summary table
        TotalStockVolume = 0 'Resets this value for the next unique ticker name
        
        YearEndValue = ws.Cells(i, 6).Value 'Record the close value for the last line of this unique ticker name
        YearlyChange = YearEndValue - YearStartValue
        ws.Cells(UniqueTickerCount + 1, 10).Value = YearlyChange 'Record the yearly change for this unique ticker name in the summary table
        
        PercentChange = YearlyChange / YearStartValue
        ws.Cells(UniqueTickerCount + 1, 11).Value = PercentChange 'Records this value in the summary table
        YearStartValue = ws.Cells(i + 1, 3).Value 'Sets this to the open of the first line for the next unique name
        
    Else
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    End If
        
Next i
'-------------------------------------------------------------------------------


'Find some global results for the set of results---------------------------------

'Finding Greatest % Increase
    Dim GreatestPercentIncrease As Double
    GreatestPercentIncrease = ws.Cells(2, 11).Value 'Starts this value with the value from the first stock
    ws.Cells(2, 17).Value = GreatestPercentIncrease
    ws.Cells(2, 16).Value = ws.Cells(2, 9).Value
    
For j = 2 To (UniqueTickerCount + 1)
    If ws.Cells(j, 11).Value > GreatestPercentIncrease Then
        GreatestPercentIncrease = ws.Cells(j, 11).Value
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
    End If
Next j

'Find Greatest % Decrease
    Dim GreatestPercentDecrease As Double
    GreatestPercentDecrease = ws.Cells(2, 11).Value 'Starts this value with the value from the first stock
    ws.Cells(3, 17).Value = GreatestPercentDecrease
    ws.Cells(3, 16).Value = ws.Cells(2, 9).Value
    
For k = 2 To (UniqueTickerCount + 1)
    If ws.Cells(k, 11).Value < GreatestPercentDecrease Then
        GreatestPercentDecrease = ws.Cells(k, 11).Value
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
    End If
Next k

'Find Greatest Total Volume
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = ws.Cells(2, 12).Value 'Starts this value with the value from the first stock
    ws.Cells(4, 17).Value = GreatestTotalVolume
    ws.Cells(4, 16).Value = ws.Cells(2, 9).Value
    
For n = 2 To (UniqueTickerCount + 1)
    If ws.Cells(n, 12).Value > GreatestTotalVolume Then
        GreatestTotalVolume = ws.Cells(n, 12).Value
        ws.Cells(4, 17).Value = GreatestTotalVolume
        ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
    End If
Next n
'--------------------------------------------------------------------------------


'Set up formatting for results--------------------------------------------------
ws.Range("1:1,O2,O3,O4").Font.Bold = True
ws.Range("I:Q").Columns.AutoFit
'Set Conditional colored fill and number format for Yearly Change results column
Dim YearlyChangeRange As Range
Set YearlyChangeRange = ws.Range("J:J")
YearlyChangeRange.FormatConditions.Delete
    Dim Condition1 As FormatCondition
    Dim Condition2 As FormatCondition
    Set Condition1 = YearlyChangeRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set Condition2 = YearlyChangeRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
        With Condition1
            .Interior.Color = RGB(0, 255, 0)
        End With
        With Condition2
            .Interior.Color = RGB(255, 0, 0)
        End With
    YearlyChangeRange.NumberFormat = "#0.00"
    ws.Range("J1").FormatConditions.Delete 'Clear conditional fill for header
    
'Set number formats
Dim PercentChangeRange As Range
Set PercentChangeRange = ws.Range("K:K")
PercentChangeRange.NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "#0.00%"
ws.Cells(3, 17).NumberFormat = "#0.00%"
ws.Cells(4, 17).NumberFormat = "#0.00E+0"
'-------------------------------------------------------------------------------------

Next ws

'Let the user know that the numbers have been crunched------------------------------------------
MsgBox ("Your results are ready.  Thank you for your patience.  You deserve a cookie!")

End Sub
