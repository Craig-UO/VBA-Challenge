Attribute VB_Name = "StockAnalyzerAllSheets"
Sub StockAnalyzerAllSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
'Set up cells for results to land in--------------------------------------------
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
'-------------------------------------------------------------------------------


'Set up list of all unique stock ticker names for column I------------------------------
'Make sure data is sorted by date and ticker name (ignores the headers)
Range("A:G").Sort Key1:=Range("B1"), Order1:=xlAscending
Range("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending

'Variable to count the number of unique ticker names to use as an index for list of names
Dim UniqueTickerCount As Integer
UniqueTickerCount = 0

'Sets the first name to get started (this will cause the header in column A to be ignored when adding new names)
Dim CurrentUniqueTicker As String
Dim CurrentTicker As String
CurrentUniqueTicker = Cells(1, 1).Value

'Adds new unique ticker names to the list (starting with the first unique name after the header)
For Each cell In Range("A:A")
    If Not IsEmpty(cell) Then
        CurrentTicker = cell.Value
        If Not (CurrentTicker = CurrentUniqueTicker) Then
            UniqueTickerCount = UniqueTickerCount + 1
            CurrentUniqueTicker = CurrentTicker
            Cells(UniqueTickerCount + 1, 9).Value = CurrentUniqueTicker
        End If
    End If
Next
'----------------------------------------------------------------------------------------


'Count the number of stock entries (rows in the file)--------------------------------------
Dim TotalRows As Double
TotalRows = -1 'Starting at -1 will mean we don't count the header in Column A
For Each cell In Range("A:A")
    If Not IsEmpty(cell) Then
    TotalRows = TotalRows + 1
    End If
Next
'-------------------------------------------------------------------------------------------


'Cycle through the list of unique ticker names and accumulate/compute data for each unique ticker name----------------
'Set variables for results to be calculated
    Dim YearlyChangeEach As Double
    Dim PercentChangeEach As Double
    Dim TotalStockVolumeEach As Double

'Set up indexes to use for range of each unique ticker name
    Dim IndexStart As Integer
    Dim IndexEnd As Integer
    
'Set up values to use to compute results
    Dim YearStartValue As Double
    Dim YearEndValue As Double
    

For i = 2 To (UniqueTickerCount + 1)
    CurrentUniqueTicker = Cells(i, 9).Value 'Sets variable to active name for each loop
    IndexStart = 0  'Resets this variable for the current loop
    IndexEnd = 0  'Resets this variable for the current loop
        
    'Determine the beginning row and earliest value for the current unique ticker name
    For j = 2 To TotalRows + 1
        If (Cells(j, 1).Value = CurrentUniqueTicker) Then
            IndexStart = j
            YearStartValue = Cells(j, 3).Value
            Exit For
        Else
        End If
    Next
   'Determine the ending row and latest value for the current unique ticker name
    For j = 2 To TotalRows + 1
        If (Cells(j, 1).Value = CurrentUniqueTicker) And (j > IndexEnd) Then
            IndexEnd = j
            YearEndValue = Cells(j, 6).Value
        End If
    Next
    
    'Compute some results and enter them into their cells
    YearlyChangeEach = YearEndValue - YearStartValue
    Cells(i, 10).Value = YearlyChangeEach
    
    PercentChangeEach = (YearEndValue - YearStartValue) / (YearStartValue)
    Cells(i, 11).Value = PercentChangeEach
    
    TotalStockVolumeEach = WorksheetFunction.Sum(Cells(IndexStart, 7), Cells(IndexEnd, 7))
    Cells(i, 12).Value = TotalStockVolumeEach
        
Next
'-------------------------------------------------------------------------------


'Find some global results for the set of results---------------------------------
'Finding Greatest % Increase
    Dim GreatestPercentIncrease As Double
    GreatestPercentIncrease = Cells(2, 11).Value 'Starts this variable with the value from the first stock
    
For i = 2 To (UniqueTickerCount + 1)
    If Cells(i, 11).Value > GreatestPercentIncrease Then
        GreatestPercentIncrease = Cells(i, 11).Value
        Cells(2, 17).Value = GreatestPercentIncrease
        Cells(2, 16).Value = Cells(i, 9).Value
    End If
Next

'Find Greatest % Decrease
    Dim GreatestPercentDecrease As Double
    GreatestPercentDecrease = Cells(2, 11).Value 'Starts this variable with the value from the first stock
    
For i = 2 To (UniqueTickerCount + 1)
    If Cells(i, 11).Value < GreatestPercentDecrease Then
        GreatestPercentDecrease = Cells(i, 11).Value
        Cells(3, 17).Value = GreatestPercentDecrease
        Cells(3, 16).Value = Cells(i, 9).Value
    End If
Next

'Find Greatest Total Volume
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = Cells(i, 12).Value
    
For i = 2 To (UniqueTickerCount + 1)
    If Cells(i, 12).Value > GreatestTotalVolume Then
        GreatestTotalVolume = Cells(i, 12).Value
        Cells(4, 17).Value = GreatestTotalVolume
        Cells(4, 16).Value = Cells(i, 9).Value
    End If
Next
'--------------------------------------------------------------------------------


'Set up formatting for results--------------------------------------------------
Range("1:1,O2,O3,O4").Font.Bold = True
Range("I:Q").Columns.AutoFit
'Set Conditional colored fill and number format for Yearly Change results column
Dim YearlyChange As Range
Set YearlyChange = Range("J:J")
YearlyChange.FormatConditions.Delete
    Dim Condition1 As FormatCondition
    Dim Condition2 As FormatCondition
    Set Condition1 = YearlyChange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set Condition2 = YearlyChange.FormatConditions.Add(xlCellValue, xlLess, "=0")
        With Condition1
            .Interior.Color = RGB(0, 255, 0)
        End With
        With Condition2
            .Interior.Color = RGB(255, 0, 0)
        End With
    YearlyChange.NumberFormat = "#0.00"
    Range("J1").FormatConditions.Delete 'Clear conditional fill for header
    
'Set number formats
Dim PercentChange As Range
Set PercentChange = Range("K:K")
PercentChange.NumberFormat = "0.00%"
Cells(2, 17).NumberFormat = "#0.00%"
Cells(3, 17).NumberFormat = "#0.00%"
Cells(4, 17).NumberFormat = "#0.00E+0"
'-------------------------------------------------------------------------------------


'Let the user know that the numbers have been crunched------------------------------------------
MsgBox ("Your results are ready.  Thank you for your patience.  You deserve a cookie!  There were " + Str(UniqueTickerCount) + " unique stocks analyzed from " + Str(TotalRows) + " data entries.")

End Sub
