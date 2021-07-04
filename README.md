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
### The advantages and disadvantages of refactoring code in general.
- Advantages:
  - The refactoring code is more readable and structured. 
  - It will be less complex of the code, which reduce the maintain cost of the script as it's more adaptive to change.
- Disadvantages:
  - Spend more time for the refactoring which may not be worth as its results.
  - While refactoring, it might create bug that not adapt the original code structure, then we will need to spend more time to fix it.

### The advantages and disadvantages of the original and refactored VBA script.
![VBA_Challenge_2017](https://github.com/rykiprince/stocks-analysis/blob/main/VBA_Challenge_2017.png) ![VBA_Challenge_2018](https://github.com/rykiprince/stocks-analysis/blob/main/VBA_Challenge_2018.png)
- By replacing the ticket index from an assigned `i` to a ca
