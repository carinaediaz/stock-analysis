# An Analysis of Stock Performances
## Overview 
##### Analyzing the data of 12 different stocks throughout 2017 and 2018 to help Steve determine which stock is best for his parents to invest in. 
## Purpose
##### Steve came to us to help him analyze how different stocks performed in 2017 and 2018, in order to make an educated recommendation to his parents on what to invest in. We utilized the daily closing prices and volumes of 12 stocks (AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR) to find the total daily volume and annual return for each. Instead of attempting to to accomplish this with Excel alone, we utilized VBA to help us create macros that automate the formating and calculating process. Automating these tasks helped us prevent human error in typing complex formulas that refer to multiple sheets within our workbook. We initially created an AllStocksAnalysis macro, that we then used as a base for our final AllStocksAnalysisRefactored macro and refactored it so that it would be more efficent. 
## Results
##### We created the AllStocksAnalysisRefactored macro to create not only the formatting of our results table, but also the values/calcuations for each of the cells within the result table. Below you will see the performance for all stocks from 2017 to 2018. 
![VBA_Challenge_Results_2017](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_Results_2017.PNG)
![VBA_Challenge_Results_2018](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_Results_2018.PNG)
##### In 2017, we see that all stocks but one (TERP) had a positive return. In 2018, we see that all stocks but ENPH and RUN had negative returns. To easily identify the positive and negative returns, we used the conditional formating for loop below. 
```
For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
```
##### Ultimately, by refactoring we were able to bring down our macro run times to below 0.15 seconds for each year (see original and refactored run times below). 
![VBA_Challenge_2017_original](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_2017_original.PNG)
![VBA_Challenge_2017](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018_original](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_2018_original.PNG)
![VBA_Challenge_2018](https://github.com/carinaediaz/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)
## Summary
##### Refactoring is essential in revisiting old code. By refactoring, we can improve the run time, debug issues, and cleaning up our code to make it easy for anyone to understand or pick up from where you left off.  Refacotring can be a useful peer-editing tool as well. It may seem like a simple task to copy previous code and refactor it until it is more efficient, but without descriptive comments, it could be difficult to follow and edit your own code, let alone someone else's. It could also become more difficult to debug an issue the more complex the code becomes. 
##### In this case, we were able to successfully refactor the macro to reduce the run time by ~83% for 2017 and ~84% for 2018 calculations. We kept the macros functionality the same, but by establishing a tickerIndex variable, and tickerVolumes, tickerStartingPrices, tickerEndingPrices arrays to use in our for loop (see code below), we drastically cut the run time. While our code runs faster, this doesn't mean there is less data. In fact, we added an index and several arrays to our code, making it more complex than the original. The more complex our coding becomes, the more descriptive comments become crutial to be able to go back and make changes without causing errors. 
```
    '1a) Create a ticker Index
    Dim tickerIndex As String
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 12
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        'End If
    
    Next i
````
