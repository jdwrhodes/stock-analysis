# Stock-Analysis

## Overview Of Project

This analysis was created to help Steve in advising his parents which stocks to invest in.

## Results

2 Macros were written to analyze a data set consisting of 12 stocks. The Total Daily Volume and ROI for these stocks was analyzed for the years of 2017 and 2018.

The first Macro used 2 nested For Loops, one array, and Start/End prices set as variables. A example of the code here: 

```vba
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        '5)Loop through rows in the data
            
            Worksheets(yearValue).Activate
            For j = 2 To RowCount
    
            '5a)Find total volume for the curent ticker
            
                If Cells(j, 1).Value = ticker Then
        
                    totalVolume = totalVolume + Cells(j, 8).Value
        
                End If
            
            '5b)Find starting price for the current ticker
            
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        
                    startingPrice = Cells(j, 6).Value
            
                End If
                
            '5c)Find ending price for the current ticker
    
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
        
                    endingPrice = Cells(j, 6).Value
        
                End If
    
            Next j
...
        
    Next i
```
This code completed the analysis in 1.64 seconds for both years, as seen in the images below.


![2017 Macro speed results](https://raw.githubusercontent.com/jdwrhodes/stock-analysis/main/Resources/Original_VBA_Challenge_2017.png "2017 Macro speed results") ![2018 Macro speed results](https://raw.githubusercontent.com/jdwrhodes/stock-analysis/main/Resources/Original_VBA_Challenge_2018.png "2018 Macro speed results")

The second Macro, which was refactored from the first, used only 1 For loop and 3 arrays. Below is a sample of that code:

```vba
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    Next i
```

This macro completed the analysis in only 0.15 seconds, a significant time savings. This is also shown in the images below.

![2017 Macro speed results](https://raw.githubusercontent.com/jdwrhodes/stock-analysis/main/Resources/VBA_Challenge_2017.png "2017 Macro speed results") ![2018 Macro speed results](https://raw.githubusercontent.com/jdwrhodes/stock-analysis/main/Resources/VBA_Challenge_2018.png "2018 Macro speed results")

## Summary
 
 The refactoring process appears to show that For Loops can be computationaly expensive. Meaning, they take more time. Using multiple arrays to store values when they are being used requently, rather than variables, seems to be faster when iterating over a large data set. However, refactoring can lead to problems if the original code was not accurately notated.
 
 In this case, refactoring the original VBA code led to a faster, more efficient program. Also, because the code was notated clearly, it was easier to deduce how the code worked and where it could be improved.
 
