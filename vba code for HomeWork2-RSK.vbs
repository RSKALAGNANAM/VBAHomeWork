Sub SummaryOfStockPerformance()

'Create variables to hold the values we need to summarize

Dim Ticker As String
Dim TotalStockVolume As Double
Dim FirstAppearanceOfTicker As Double
Dim LastAppearanceOfTicker As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

'Create a counter for each unique Ticker Row to write the results
Dim UniqueTickerCounter As Integer

'Loop through all Worksheets

For Each ws In Worksheets
    'Initialize TotalStockVolume as 0

    TotalStockVolume = 0

    'Initialize the value of UniqueTickerCounter as 2 because results begin from Row 2

    UniqueTickerCounter = 2

    'Determine the last Row in the current Worksheet

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Add 1 to the LastRow to identify the first Blank Row
    LastRow = LastRow + 1

    'Determine the Last Column in the Current Worksheet

    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    'Insert Headers for Summaries

    ws.Cells(1, LastColumn + 2).Value = "Ticker"
    ws.Cells(1, LastColumn + 3).Value = "Opening Price"
    ws.Range("J1").Columns.AutoFit
    ws.Cells(1, LastColumn + 4).Value = "Closing Price"
    ws.Range("K1").Columns.AutoFit
    ws.Cells(1, LastColumn + 5).Value = "Yearly Change"
    ws.Range("L1").Columns.AutoFit
    ws.Cells(1, LastColumn + 6).Value = "Percent Change"
    ws.Range("M1").Columns.AutoFit
    ws.Cells(1, LastColumn + 7).Value = "Total Volume"

    'Go through the current sheet till the Last Row

    For i = 2 To LastRow
     
       'Check if we are still with the same Ticker

       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          'Register LastTimeAppearance of a Ticker
          LastAppearanceOfTicker = i
          'Set the value of Ticker to the current Ticker
          Ticker = ws.Cells(i, 1).Value
          'Determine the Total Stock Volume for the current Ticker
          TotalStockVolume = TotalStockVolume + ws.Cells(i, LastColumn).Value
          'Set the opening and closing values of the current Ticker if there was trading
          If TotalStockVolume > 0 Then
             OpeningPrice = ws.Cells(FirstAppearanceOfTicker, 3).Value
             ClosingPrice = ws.Cells(LastAppearanceOfTicker, 6).Value
             YearlyChange = ClosingPrice - OpeningPrice
             PercentChange = YearlyChange / OpeningPrice
          Else
             OpeningPrice = 0
             ClosingPrice = 0
             YearlyChange = 0
             PercentChange = 0
          End If
          'ClosingPrice = ws.Cells(i, 6).Value
          'Write the values to the Summary Table
          ws.Cells(UniqueTickerCounter, LastColumn + 2).Value = Ticker
          ws.Cells(UniqueTickerCounter, LastColumn + 3).Value = OpeningPrice
          ws.Cells(UniqueTickerCounter, LastColumn + 4).Value = ClosingPrice
          ws.Cells(UniqueTickerCounter, LastColumn + 5).Value = YearlyChange
          'Choose Green or Red to Fill Cell Based on Positive Change or Negative Change
          'No fill if Change is "zero"
          If YearlyChange > 0 Then
             ws.Cells(UniqueTickerCounter, LastColumn + 5).Interior.ColorIndex = 4
          ElseIf YearlyChange < 0 Then
             ws.Cells(UniqueTickerCounter, LastColumn + 5).Interior.ColorIndex = 3
          End If
          ws.Cells(UniqueTickerCounter, LastColumn + 6).Value = PercentChange
          'Convert PercentChange display to Percent
          ws.Cells(UniqueTickerCounter, LastColumn + 6).NumberFormat = "0.00%"
          ws.Cells(UniqueTickerCounter, LastColumn + 7).Value = TotalStockVolume
          'Reset TotalStockVolume to 0
          TotalStockVolume = 0
          'Reset FirstAppearancesOfTicker to 0
          FirstAppearanceOfTicker = 0
          'Increment UniqueTickerCounter by 1 to record results of the next Ticker Symbol
          UniqueTickerCounter = UniqueTickerCounter + 1

       Else
          'Continue tracking first appearances of the same Ticker
          If ws.Cells(i, 3).Value > 0 Then
             If FirstAppearanceOfTicker = 0 Then
                FirstAppearanceOfTicker = i
             End If
          End If
          'Continue calculating the Total Stock Volume for the current Ticker
          TotalStockVolume = TotalStockVolume + ws.Cells(i, LastColumn).Value
      
       End If
   
    Next i
    
    'Autofit the column widths
    
    'ws.Range("I1:N1").Columns.AutoFit
    
  Next ws

  'Now add the code for "Hard" part of the assignment

  For Each ws In Worksheets
     'Identify the Last Column
    
      LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
     'Identify the Last Row in the Range where we wrote the results
     'For this we need to define the First Column of the Results
   
      Dim FirstColumn As Integer
      FirstColumn = LastColumn - 5 'hard coded 5 because we have 6 columns
    
      'Now, identify the Last Row in the Last Column
      With ActiveSheet
         LastRow = .Cells(.Rows.Count, "N").End(xlUp).Row
      End With
    
      'Write the Headers for the Summary Part
      ws.Cells(1, LastColumn + 4).Value = "Ticker"
      ws.Cells(1, LastColumn + 5).Value = "Value"
      ws.Cells(2, LastColumn + 3).Value = "Greatest Increase"
      ws.Cells(3, LastColumn + 3).Value = "Greatest Decrease"
      ws.Cells(4, LastColumn + 3).Value = "Greatest Total Volume"
      ws.Range("Q4").Columns.AutoFit
    
    
      'Estimate the required Statistics
      'Find largest Percent increase and associated Ticker
      Dim TickerLargestIncrease As String
      Dim TickerLargestDecrease As String
      Dim TickerGreatestVolume As String
      Dim LargestIncrease As Double
      Dim LargestDecrease As Double
      Dim GreatestTotalVolume As Double
    
      'Initialize values of LargestIncrease,LargestDecrease and GreatesTotalVolume to 0
      LargestIncrease = 0
      LargestDecrease = 0
      GreatestTotalVolume = 0
    
      For i = 2 To LastRow
         If ws.Cells(i, LastColumn - 1).Value > LargestIncrease Then
            LargestIncrease = ws.Cells(i, LastColumn - 1).Value
            TickerLargestIncrease = ws.Cells(i, LastColumn - 5).Value
         End If
         If ws.Cells(i, LastColumn - 1).Value < LargestDecrease Then
            LargestDecrease = ws.Cells(i, LastColumn - 1).Value
            TickerLargestDecrease = ws.Cells(i, LastColumn - 5).Value
         End If
         If ws.Cells(i, LastColumn).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(i, LastColumn).Value
            TickerGreatestVolume = ws.Cells(i, LastColumn - 5).Value
         End If
      Next i
      
     'Write the values in the Summary Table
      ws.Cells(2, LastColumn + 4).Value = TickerLargestIncrease
      ws.Cells(2, LastColumn + 5).Value = LargestIncrease
      ws.Cells(2, LastColumn + 5).NumberFormat = "0.00%"
      ws.Cells(3, LastColumn + 4).Value = TickerLargestDecrease
      ws.Cells(3, LastColumn + 5).Value = LargestDecrease
      ws.Cells(3, LastColumn + 5).NumberFormat = "0.00%"
      ws.Cells(4, LastColumn + 4).Value = TickerGreatestVolume
      ws.Cells(4, LastColumn + 5).Value = GreatestTotalVolume
      ws.Cells(4, LastColumn + 5).NumberFormat = "0"
      
      ws.Range("R1").Columns.AutoFit
      ws.Range("S4").Columns.AutoFit
     
  Next ws
    
    

End Sub

