# An Analysis of Energy Stocks Using VBA

## Overview of Stock Analysis

### Purpose
This analysis of Energy industry stocks is to statistically visualize company returns to make better investment decisions.   The financical metrics used to compare performance are Year over Year Return and Total Daily Volume. The basket of stocks pre selected for this analysis include: AY,CSIQ,DQ,ENPH,FSLR,HASI,JKS,RUN,SEDG,SPWR,TERP,and VSLR. The datset researched is focused on year's: 2017 & 2018; as well as the daily stock prices at Open/Close, High/Low & Adjusted Close. Implementing Macro/Subroutine code will yield the best performing stocks in years 2017 and 2018 respectively. 


## Results and Analysis 

The Subroutine created in the module allows to anyalze the same dataset by specific year (e.g. sheet 2018). In order to determine which stocks are the best performing Year over Year the following Subroutine is refactored to include sheet 2017 and sheet 2018:

![Code for tickerIndex, output arrays & loop](https://github.com/MoKmo176/Stock_Analysis/blob/67e8faf1e1fac11e103a699f553bc2ee260901c2/CodeResources/Code%201_2.png)


The tickerIndex is increased in a conditional loop to avoid a syntax error (Div/0).Furthermore in order to set parameters for our conditionals the followng code is implemented:

![Increase ticker volume formula via StackOverflow, If Then(s), loop through output arrays ](https://github.com/MoKmo176/Stock_Analysis/blob/67e8faf1e1fac11e103a699f553bc2ee260901c2/CodeResources/Code%203_4.png)



---

The following Macro calculates the annual return of a basket of 12 stocks.  

- The results yield positive returns for 11 out of 12 stock: 

	![2017 Returns](https://github.com/MoKmo176/Stock_Analysis/blob/b36b90eeabf9f19aaaa689f9d388bd4de5384223/Resources/VBA_Challenge_2017.png)

- The results of the same stocks yield 10 out 12 negative returns in the year 2018.

	![2018 Returns](https://github.com/MoKmo176/Stock_Analysis/blob/67e8faf1e1fac11e103a699f553bc2ee260901c2/Resources/VBA_Challenge_2018.png)



## Summary


1. The advantages of refactoring code in general are that we can statiscally prove which stocks outperform Year over Year. In this Subroutine, we have a numbers-based case that ENPH and RUN had positive Year over Year returns. We can also determine that DQ would have underperformed ENPH Year over Year, given that DQ had the best performing Return in 2017.  

The disadvantages include that we don't have Average Yearly Returns and other stock metric analysis. The original code is a precusor to the refactor anaylsis, as well. Analyzing the data within smaller segments can also be useful. 

2. The advantages of the original and refactored allows to analyze the dataset microscopiclly. We can disect the data by Open/Close prices, High/Low prices as well as adjusted closing prices to hyperanalyze institutional before/after hours market trading. Another advantage would include visualizing a specific Years data through a pivot table or graph without having to filter too much data. 

The disadvantages include having syntax errors when running code due to refactoring errors in the data set. Delineating patterns becomes harder if there was a large market shifting event that resulted in underperforming stocks in 2 different calendar years. Analysing the data in the refactored analysis of stock performance gives us the data we need to determine market fluctuations and outperforming stocks. 
