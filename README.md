# Module-2-Challenge
VBA_Challenge
# 1. Overview of Project
In this challenge we are doig a Stock Analysis for the follwing stocks with tickers  : AY, CSIG, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR.Steve's parents wanted some information on before investing. we used VBA with an excel file that had data for each ticker that we used to perform analysis . in this project we needed to calculate the Total Daily Value and the annual return for each of the 12 stocks in this data . so we can determinate which stocks in this data performed the best to worst.

# 2. Results
 # Explain the purpose of this analysis
the purpose of this code was to loop through all the data for the years (2018 and 2017) and get the infomation needed (the total daily volume of each stock and annual return ).in this code we have created  created a 4 different arrays; tickers (for each stock) , tickerVolumes ,tickerStartingPrices and tickerEndingPrices 
and we used tickerIndex to match the arrays with the tickers .
# the code :

Sub AllStocksAnalysisRefactored()
Dim startTime As Single
Dim endTime  As Single

yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

'Format the output sheet on All Stocks Analysis worksheet
Worksheets("All Stocks Analysis").Activate

Range("A1").Value = "All Stocks (" + yearValue + ")"
'Create a header row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'Initialize array of all tickers
Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'Activate data worksheet
Worksheets(yearValue).Activate

'Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'1a) Create a ticker Index
Dim tickerIndex As Single
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
tickerVolumes(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.
For j = 2 To RowCount

    '3a) Increase volume for current ticker
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
     tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
        
    'End If
End If

    '3c) check if the current row is the last row with the selected ticker
     'If the next row’s ticker doesn’t match, increase the tickerIndex.
    'If  Then
If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        

        '3d Increase the tickerIndex.
    tickerIndex = tickerIndex + 1
        
    'End If
    End If
    
Next j

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
    
Next i

'Formatting
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.FontStyle = "Bold"
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

dataRowStart = 4
dataRowEnd = 15

For i = dataRowStart To dataRowEnd
    
    If Cells(i, 3) > 0 Then
        
        Cells(i, 3).Interior.Color = vbGreen
        
    Else
    
        Cells(i, 3).Interior.Color = vbRed
        
    End If
    
Next i

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
# 3a) Initialize variables for starting price and ending price
Dim startingPrice As Double
Dim endingPrice As Double
# 3b) Activate data worksheet
Worksheets(yearValue).Activate
# 3c) Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
# 4) Loop through tickers
For i = 0 To 11
ticker = tickers(i)
TotalVolume = 0
Worksheets(yearValue).Activate
# 5) loop through rows in the data

For j = 2 To RowCount

# 5a) Get total volume for current ticker

If Cells(j, 2).Value = ticker Then

    'increase totalVolume by the value in the current row
    TotalVolume = TotalVolume + Cells(j, 8).Value
End If

 # 5b) get starting price for current ticker

If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
    'set starting price
    startingPrice = Cells(j, 6).Value

End If

 # 5c) get ending price for current ticker
    
    If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
    'set ending price
    endingPrice = Cells(j, 6).Value

End If

Next j
# 6) Output data for current ticker

Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = TotalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
Next i


# Elapsed Time to Run for each year (2017 & 2018)
 2018

![image](https://user-images.githubusercontent.com/93894919/145730080-881ed672-2801-4595-bf71-723b407e2605.png)
2017
![image](https://user-images.githubusercontent.com/93894919/145730127-a95d6b5a-f2bd-4e36-8693-7a0374be6f17.png)

# Summary 

Refactoring the code in small steps we will make us understad the code better and makinge tiny changes in your program will make your code better and leaves the application in a working state , also we can see easily the logical errors .
and VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source
there are some disadvantages of refactoring the code for exemple in the long procedure may contain the same line of the code in diffrenet locations ,you can change the logic to eliminate the duplicate lines.also Refactoring process can affect the testing outcomes.
How do these pros and cons apply to refactoring the original VBA script?
for my opinion A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.



