Attribute VB_Name = "Module1"
Sub processStockData():
'This one has the whole data
'Define and assign worksheets to variables

For Each ws In Worksheets
'Dim sheet1, sheet2, sheet3, sheet4, sheet5, sheet6 As Worksheet
Dim WorksheetName As String
WorksheetName = ws.Name

'Set sheet1 = Worksheets("A")
'Set sheet2 = Worksheets("B")
'Set sheet3 = Worksheets("C")
'Set sheet4 = Worksheets("D")
'Set sheet5 = Worksheets("E")
'Set sheet6 = Worksheets("F")

'Populate the column header in Sheet
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'declaration of variable
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'get the last row in the spreadsheet
totalStockVolume = 0   'holds total stock volume for each ticker
tickerRowCounter = 2   'start at the first row with data for the column
openPriceRow = 2       'starting row that hold the opening price for the ticker
lastrowPercent = Cells(Rows.Count, 11).End(xlUp).Row 'lastrow in Percent Change column
lastRowTotalStockVolume = Cells(Rows.Count, 12).End(xlUp).Row 'last row in total stock column
Dim maxPercentIncrease, maxPercentDecrease As Double
Dim openingPrice, endingPrice As Double
'Dim sheet1, sheet2, sheet3 As Worksheet

'Set sheet1 = Worksheets("A")
'Set sheet2 = Worsheets("2019")
'Set sheet3 = Worksheets("2020")


'MsgBox (lastrow)

For Row = 2 To lastRow:
    'Check if there is a change in the ticker symbol starting the value in row 2
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
         
     'If tickets are not equal, record value of ticker in row corresponding to tickerRowCounter in Column I
   
     ws.Cells(tickerRowCounter, 9) = ws.Cells(Row, 1).Value
   
     'Calculate the total stock volume per ticker
     ws.Cells(tickerRowCounter, 12) = totalStockVolume + ws.Cells(Row, 7).Value
     
     'set the sending and opening Price Values
     endingPrice = ws.Cells(Row, 6).Value
     openingPrice = ws.Cells(openPriceRow, 3).Value
     
     'MsgBox ("Opening Price  " & openPrice)
     'MsgBox ("ending Price  " & endingPrice & " Row: " & row)
     
     'Calculate and record the Yearly Change
     
     YearlyChange = endingPrice - openingPrice
     ws.Cells(tickerRowCounter, 10) = YearlyChange
     
     'If Yearly Change positive  cell background green, if negative background red
      If YearlyChange < 0 Then
         ws.Cells(tickerRowCounter, 10).Interior.ColorIndex = 3 ' Red
        Else
         ws.Cells(tickerRowCounter, 10).Interior.ColorIndex = 4 'Green
     End If
     
     'Calcuate Percent of Change per ticker, record in column 11 (K) and format values
     ws.Cells(tickerRowCounter, 11) = YearlyChange / openingPrice
     ws.Cells(tickerRowCounter, 11).NumberFormat = "#.##%"
     
     totalStockVolume = 0   'reset total stock volume to calcuate total for next ticker
     tickerRowCounter = tickerRowCounter + 1    'increment ticker counter
     openPriceRow = Row + 1 'get the next opening price
    
    Else
      
       totalStockVolume = totalStockVolume + Cells(Row, 7).Value ' running total for current ticker
       
    End If

Next Row
    
    ' Get the greatest percentage of increase
     maxPrecentIncrease = WorksheetFunction.max(ws.Range("K2:K" & lastrowPercent)) 'Find greatest value in Column and assign to MaxPrecentage variable
     maxPrecentIncreaseIndex = WorksheetFunction.Match(maxPrecentIncrease, ws.Range("K2:K" & lastrowPercent), 0) 'Find the Index (row) of the greatest value
     ws.Cells(2, 17).Value = maxPrecentIncrease  'record the value of MaxPrecentage Increase
     ws.Cells(2, 17).NumberFormat = "#.##%"      'format the cell
     ws.Cells(2, 16).Value = Cells(maxPrecentIncreaseIndex + 1, 9) 'Record the ticker associated with the MaxPrecentage Increase
     
     
     maxPrecentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & lastrowPercent)) 'Find decrease value in Column and assign to MaxPrecentage variable
     maxPrecentDecreaseIndex = WorksheetFunction.Match(maxPrecentDecrease, ws.Range("K2:K" & lastrowPercent), 0) 'Find the Index (row) of the greatest value
     ws.Cells(3, 17).Value = maxPrecentDecrease  'record the value of MaxPrecentage Decrease
     ws.Cells(3, 17).NumberFormat = "#.##%"      'format the cell
     ws.Cells(3, 16).Value = ws.Cells(maxPrecentDecreaseIndex + 1, 9)  'Record the ticker associated with the MaxPrecentage Increase
     
     'Greatest Total Volume
     
     maxTotalStockVolume = WorksheetFunction.max(ws.Range("L2:L" & lastRowTotalStockVolume)) 'Find Max Total Stock Volume in Column and assign to MaxPrecentage variable
     maxTotalStockVolumeIndex = WorksheetFunction.Match(maxTotalStockVolume, ws.Range("L2:L" & lastRowTotalStockVolume), 0) 'Find the Index (row) Max Total Stock Volume
     ws.Cells(4, 17).Value = maxTotalStockVolume  'record the value of Max Total Stock Volume
     ws.Cells(4, 16).Value = Cells(maxTotalStockVolumeIndex + 1, 9)  'Record the ticker associated with the Max Total Volume
     
     ws.Range("A:Q").Columns.AutoFit 'autofit size the columns
 'Exit For
Next ws
End Sub

