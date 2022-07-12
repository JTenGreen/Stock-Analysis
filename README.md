# Stock-Analysis
Purponse:
The purpose of this challenge was to compare green stock data from 2017 and 2018 to see which option would be the best choice for Steve's parents who want to invest in green stocks or sustainable energy stocks.
Results:
By refactoring the code we could reduce a .9 second time to run down to .1 second run time by optimizing the code to run in a more efficient way as seen below.

 '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 12
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
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
         End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(1 + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        Next i
        
        ![image](https://user-images.githubusercontent.com/106503121/178381190-43b43b9d-a8b1-4f4e-94fb-a0a207c9ee10.png)
        
        ![image](https://user-images.githubusercontent.com/106503121/178381232-94805972-cafa-4379-9fe9-ce18a8286d62.png)


Summary:
As a result of the analysis we can determine that DQ is not a good stock to invest in as of 2018 it was down by 62%, while ENPH and RUN are up 81.9% and 84% respectively for 2018 and they were the only 2 stocks during 2018 to now go down in value.

The advantage of refactoring code is that we can clean it up and organize it better and seen from a reader point of view we can even find flaws that the writer didn't notice, a fresh set of eyes can improve code and optimize it.
In this case by refavtoring the code we were able to bring down the time by .8 seconds.


