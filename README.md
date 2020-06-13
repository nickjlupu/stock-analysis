# stock-analysis

# challenge

In this challenge the class was tasked with refactoring the code that was developed during the modules.  The initial code had to loop through all of the data for each ticker.  By refactoring, I have simplified the code to only loop through the data once and store all of the required information into arrays.  These arrays are in contrast to individual variables that would be cumbersome to set up and maintain.  The data stored in the arrays are then output to the worksheet by looping through the elements of each array. 

## comments

This code assumes that any further data by year would be introduced to the excel sheet as a separate tab with the title of the tab being that year in question.  Furthermore, if additional tickers are to be analyzed, the code would need to be refactored further in order to expand the arrays and define their ticker symbols.

### below is the refactored code found in the attached Excel book:

Sub AllStocksAnalysis()
    
    'get user input for year to run analysis
    Dim yearValue As String
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


    'initialize array of all tickers we wish to evaluate
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
    
     
    
    ' initialize arrays for starting price and ending price (add array for volume)
    Dim startingPrice(12) As Single
    Dim endingPrice(12) As Single
    Dim volume(12) As Long
    For n = 0 To 11
    volume(n) = 0
    Next n
    
    
    'create ticker index variable & initialize
    
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    ' get the number of rows to loop over in correct worksheet
    Worksheets(yearValue).Activate
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through rows in the data
    For i = 2 To RowCount
        Worksheets(yearValue).Activate
        ticker = tickers(tickerIndex)
          
        ' get total volume for current ticker
        If Cells(i, 1).Value = ticker Then
            volume(tickerIndex) = volume(tickerIndex) + Cells(i, 8).Value
        End If
            
        ' get starting price for current ticker
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            startingPrice(tickerIndex) = Cells(i, 6).Value
        End If
            
        ' get ending price for current ticker
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            endingPrice(tickerIndex) = Cells(i, 6).Value
        End If
   
        ' increase ticker index if next row ticker is not equal to current row ticker
        If Cells(i + 1, 1).Value <> ticker Then
        tickerIndex = tickerIndex + 1
        End If
   
   Next i
    
    
    ' output data
    For j = 0 To 11
   
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + j, 1).Value = tickers(j)
        Cells(4 + j, 2).Value = volume(j)
        Cells(4 + j, 3).Value = endingPrice(j) / startingPrice(j) - 1
   
   Next j
   
   
   ' format cells
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    ' conditional format returns for clarity
    dataRowStart = 4
    dataRowEnd = 15
    For k = dataRowStart To dataRowEnd
        If Cells(k, 3) > 0 Then
            'Color the cell green
            Cells(k, 3).Interior.Color = vbGreen
        ElseIf Cells(k, 3) < 0 Then
            'Color the cell red
            Cells(k, 3).Interior.Color = vbRed
        Else
            'Clear the cell color
            Cells(k, 3).Interior.Color = xlNone
        End If

    Next k
    
End Sub
