# Stock-Analysis
 The purpose of this project was to refactor the code to create a loop for the VBA code that runs faster than the original code.
 New Code: 
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
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    Worksheets(yearValue).Activate
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
       
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         'End if
         End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
                 
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
            
        
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

 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/93847102/148659963-e9fdbe09-0d99-43c3-a5ad-0733d8922cf8.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/93847102/148659966-8a9a9932-6ae0-490c-b6f2-a54c2fabe92a.png)

Old Code:
Sub AllStocksAnalysis()
        Dim startTime As Single
        Dim endTime As Single
        
        yearValue = InputBox("What year would you like to run the analysis on?")
    
            startTime = Timer

    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks ((" + yearValue + ")"
    
    'Create a Header Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize an array of all tickers
    Dim tickers(11) As String
    
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
    
    
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
   'Loop through the tickers
   For i = 0 To 11
   ticker = tickers(i)
   totalVolume = 0

   'Loop through rows in the data
   Worksheets(yearValue).Activate
     For j = 2 To RowCount
     
   
   'Find the total volume for the current ticker
   If Cells(j, 1).Value = ticker Then
   
    totalVolume = totalVolume + Cells(j, 8).Value
   
   End If
   
   If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
   
    startingPrice = Cells(j, 6).Value
    
  End If
  
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        endingPrice = Cells(j, 6).Value
        
        
    End If
    Next j
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
    
        endTime = Timer
        MsgBox "This Code ran in " & (endTime - startTime) & " seconds for the year" & (yearValue)
    
End Sub

The advantage of refactoring code is to make it more efficienct.
![original 2017 analysis](https://user-images.githubusercontent.com/93847102/148660173-0c978b89-20d2-4fc0-8b09-b4c59a2a10db.JPG)
![2018 Original Analysis](https://user-images.githubusercontent.com/93847102/148660176-70aaad16-cc0d-4ce2-8673-175c94ba08ec.JPG)

The purpose of refactoring code is to make it run more efficiently.
