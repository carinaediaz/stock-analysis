#An Analysis of 2017 vs 2018 Stock Performances
##Overview 
#####Analyzing the data of 12 different stocks throughout 2017 and 2018 to help Steve determine which stock is best for his parents to invest in. 
##Purpose
#####Steve came to us to help him analyze how different stocks performed in 2017 and 2018, in order to make an educated recommendation to his parents on what to invest in. We utilized the daily closing prices and volumes of 12 stocks (AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR) to find the total daily volume and annual return for each. Instead of attempting to to accomplish this with Excel alone, we utilized VBA to help us create macros that automate the formating and calculating process. Automating these tasks helped us prevent human error in typing complex formulas that refer to multiple sheets within our workbook. We initially created an AllStocksAnalysis macro, that we then used as a base for our final AllStocksAnalysisRefacoted macro and refractored it so that it would be more efficent. 
##Results
#####We created the AllStocksAnalysisRefactored macro to create not only the formatting of our results table, but also the values/calcuations for each of the cells within the result table. Below you will see the performance for all stocks from 2017 to 2018. 
![VBA_Challenge_Results_2017](https://github.com/carinaediaz/stock-analysis/blob/main/VBA_Challenge_Results_2017.PNG)
![VBA_Challenge_Results_2018](https://github.com/carinaediaz/stock-analysis/blob/main/VBA_Challenge_2018.PNG)
In 2017, we see that all stocks but one (TERP) had a positive return. In 2018, we see that all stocks but ENPH and RUN had negative returns. To easily identify the positive and negative returns, we used the conditional formating for loop below. 
```
For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
```
#####Ultimately, by refactoring we were able to bring down our macro run times to below 0.15 seconds for each year (see screenshots below). 
![VBA_Challenge_2017](https://github.com/carinaediaz/stock-analysis/blob/main/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/carinaediaz/stock-analysis/blob/main/VBA_Challenge_2018.PNG)
##Summary
#####
