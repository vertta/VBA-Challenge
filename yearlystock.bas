Attribute VB_Name = "Module11"
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
 'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
 'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.


Sub yearlystock():

'declaration of variable
lastrow = Cells(Rows.Count, 1).End(xlUp).row 'get the last row in the spreadsheet
totalStockVolume = 0   'holds total stock volume for each ticker
tickerRowCounter = 2   'start at the first row with data for the column
openPriceRow = 2       'starting row that hold the opening price for the ticker
lastrowPercent = Cells(Rows.Count, 11).End(xlUp).row 'lastrow in Percent Change column
lastRowTotalStockVolume = Cells(Rows.Count, 12).End(xlUp).row 'last row in total stock column

Dim maxPercentIncrease, maxPercentDecrease As Double
Dim openingPrice, endingPrice As Double
Dim sheet1, sheet2, sheet3 As Worksheet

Set sheet1 = Worksheets("2018")
Set sheet2 = Worsheets("2019")
Set sheet3 = Worksheets("2020")


'MsgBox (lastrow)

For row = 2 To lastrow:
    'Check if there is a change in the ticker symbol starting the value in row 2
    If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
         
     'If tickets are not equal, record value of ticker in row corresponding to tickerRowCounter in Column I
   
     Cells(tickerRowCounter, 9) = Cells(row, 1).Value
   
     'Calculate the total stock volume per ticker
     Cells(tickerRowCounter, 12) = totalStockVolume + Cells(row, 7).Value
     
     'set the sending and opening Price Values
     endingPrice = Cells(row, 6).Value
     openingPrice = Cells(openPriceRow, 3).Value
     
     'MsgBox ("Opening Price  " & openPrice)
     'MsgBox ("ending Price  " & endingPrice & " Row: " & row)
     
     'Calculate and record the Yearly Change
     
     YearlyChange = endingPrice - openingPrice
     Cells(tickerRowCounter, 10) = YearlyChange
     
     'If Yearly Change positive  cell background green, if negative background red
      If YearlyChange < 0 Then
         Cells(tickerRowCounter, 10).Interior.ColorIndex = 3 ' Red
        Else
         Cells(tickerRowCounter, 10).Interior.ColorIndex = 4 'Green
     End If
     
     'Calcuate Percent of Change per ticker, record in column 11 (K) and format values
     Cells(tickerRowCounter, 11) = YearlyChange / openingPrice
     Cells(tickerRowCounter, 11).NumberFormat = "#.##%"
     
     totalStockVolume = 0   'reset total stock volume to calcuate total for next ticker
     tickerRowCounter = tickerRowCounter + 1    'increment ticker counter
     openPriceRow = row + 1 'get the next opening price
    
    Else
      
       totalStockVolume = totalStockVolume + Cells(row, 7).Value ' running total for current ticker
       
    End If

Next row
    ' Get the greatest percentage of increase
     maxPrecentIncrease = WorksheetFunction.max(sheet1.Range("K2:K" & lastrowPercent)) 'Find greatest value in Column and assign to MaxPrecentage variable
     maxPrecentIncreaseIndex = WorksheetFunction.Match(maxPrecentIncrease, sheet1.Range("K2:K" & lastrowPercent), 0) 'Find the Index (row) of the greatest value
     Cells(2, 17).Value = maxPrecentIncrease  'record the value of MaxPrecentage Increase
     Cells(2, 17).NumberFormat = "#.##%"      'format the cell
     Cells(2, 16).Value = Cells(maxPrecentIncreaseIndex + 1, 9) 'Record the ticker associated with the MaxPrecentage Increase
     
     
     maxPrecentDecrease = WorksheetFunction.Min(sheet1.Range("K2:K" & lastrowPercent)) 'Find decrease value in Column and assign to MaxPrecentage variable
     maxPrecentDecreaseIndex = WorksheetFunction.Match(maxPrecentDecrease, sheet1.Range("K2:K" & lastrowPercent), 0) 'Find the Index (row) of the greatest value
     Cells(3, 17).Value = maxPrecentDecrease  'record the value of MaxPrecentage Decrease
     Cells(3, 17).NumberFormat = "#.##%"      'format the cell
     Cells(3, 16).Value = Cells(maxPrecentDecreaseIndex + 1, 9)  'Record the ticker associated with the MaxPrecentage Increase
     
     'Greatest Total Volume
     
     maxTotalStockVolume = WorksheetFunction.max(sheet1.Range("L2:L" & lastRowTotalStockVolume)) 'Find Max Total Stock Volume in Column and assign to MaxPrecentage variable
     maxTotalStockVolumeIndex = WorksheetFunction.Match(maxTotalStockVolume, sheet1.Range("L2:L" & lastRowTotalStockVolume), 0) 'Find the Index (row) Max Total Stock Volume
     Cells(4, 17).Value = maxTotalStockVolume  'record the value of Max Total Stock Volume
     Cells(4, 16).Value = Cells(maxTotalStockVolumeIndex + 1, 9)  'Record the ticker associated with the Max Total Volume
End Sub
