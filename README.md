# An Analysis of Stocks Performing
## Overview of Project
### Using VBA script to analyze stock trend and performance in between 2017 and 2018. Then, we turns out a script that can uncover how actively the stocks were traded and how they returns at the end of the year.
To create this VBA script, we can help Steve and his parent to have a view of how these stock perform, and find out a better choice for their investment.
## Results
### Comparing the Total Daily Volume for all stocks
![Total_Volume_Comparison_Chart](https://user-images.githubusercontent.com/66225050/124364141-62bd2c00-dbf4-11eb-9b32-32f71197e1d3.png)
Out of all 12 stocks of this data sample, there are 5 stocks reduced daily trading volume, and 7 of them increased. Indicating how often the stocks get traded, the total volume among **$DQ** and **$HASI** are relatively smaller than other stocks even they are increase, which  And, with a significant higher volume than others, **$ENPH** and **$RUN** have a dramatic increase volume from 2017 to 2018.  
### Comparing the Return for all stocks
![2017   2018 Comparison](https://user-images.githubusercontent.com/66225050/124393459-7bd6e300-dcaf-11eb-9c2d-fe65f31b0a7b.png)

Overall, most of the stocks return positively in 2017 but negatively in 2018. Only **$ENPH** and **$RUN** keep being positive from 2017 to 2018. The return of **$DQ** was dropping from 199% to (-62.6%).
### Conclusion
The stock that Steve's parent really interested in, **$DQ**, has a comparably low trading volume during 2017 and 2018. In other words, it was not oftenly traded. It also has a very negative return dropping from positive to negative. Showing on the contrast, **$ENPH** and **$RUN** were traded with a way more higher increasing volume, and also kept having positive return in both 2017 and 2018.
## Summary
### What are the advantages or disadvantages of refactoring code?
- Advantages:
  - The refactoring code is more readable and structured. 
  - It will be less complex of the code, which reduce the maintain cost of the script as it's more adaptive to change.
- Disadvantages:
  - Spend more time for the refactoring which may not be worth as its results. As while refactoring, it might create bug that not adapt the original code structure, then we will need to spend more time to fix it.

### How do these pros and cons apply to refactoring the original VBA script?
**Original script and code timer**
  ```
  Sheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
                        
        Next j
  Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
```
  ![VBA_Challenge_2017(original)](https://github.com/rykiprince/stocks-analysis/blob/main/Resources/VBA_Challenge_2017(original).png) ![VBA_Challenge_2018(original)](https://github.com/rykiprince/stocks-analysis/blob/main/Resources/VBA_Challenge_2018(original).png)

**Refactored Script and code timer**
```
 RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    tickerIndex = 0

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    For i = 2 To RowCount
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If        
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            tickerIndex = tickerIndex + 1
            
        End If      
    
    Next i
    
    For i = 0 To 11
        
        Sheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
  ![VBA_Challenge_2017](https://github.com/rykiprince/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018](https://github.com/rykiprince/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)
- The refactored code is now cleaner and more readable and organized with new more clear named variables and less complex nested statements.
- The refactored code eliminated the nested loop improving the efficiency of the script. When Tom wants to expand this code to the entire stock market, this refactored code will help with less complex loop script.
- Comparing the code timers of the original code and the refactored one, the original takes about 0.3s for both years, and the refactored taks about 0.05s. it's obviously that the code performs well with a shorter time - more efficient. This results worth the effort we put on refactoring
