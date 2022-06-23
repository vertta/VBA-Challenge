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

Dim openingPrice As Double
Dim endingPrice As Double

'MsgBox (lastrow)

For row = 2 To lastrow
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
     YearlyChange = endingPrice - openingPrice
     Cells(tickerRowCounter, 10) = YearlyChange   'Calculate and record the Yearly Change
     
     'If Yearly Change positive  cell background green, if negative background re
     If YearlyChange < 0 Then
         Cells(tickerRowCounter, 10).Interior.ColorIndex = 3 ' Red
        Else
         Cells(tickerRowCounter, 10).Interior.ColorIndex = 4 'Green
     End If
     
     Cells(tickerRowCounter, 11) = YearlyChange / openingPrice
     Cells(tickerRowCounter, 11).NumberFormat = "#.##%"
     
     totalStockVolume = 0   'reset total stock volume to calcuate total for next ticker
     tickerRowCounter = tickerRowCounter + 1    'increment ticker counter
     openPriceRow = row + 1 'get the next opening price
  
    Else
      
       totalStockVolume = totalStockVolume + Cells(row, 7).Value ' running total for current ticker
       
    End If
        
      
     
'MsgBox (Cells(row).Value & " next value are equal'Cells(row + 1).Value)
Next row


