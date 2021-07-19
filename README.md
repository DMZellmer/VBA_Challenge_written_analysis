**# VBA Challenge Written Analysis**
 
**##Overview of Project**

The purpose of this analysis was to edit, or refactor, the VBA code used to read and output various stock tickers into an analysis worksheet for easy comparison. This refactoring was undertaken in order to speed up the run time when analyzing larger datasets. 

##**Results**

 ###**Hot Stock: ENPH**

Looking at the stock data for 2017 and 2018 for these 12 tickers, it's clear that 2017 was a good year for almost all and 2018 was a bad year for almost all. Except ENPH. This stock was in the green both years, with exceptional performance in 2017 and 2018.  
    

###**High-Performing ENPH Stock**

![ENPH 2017](https://user-images.githubusercontent.com/86337475/126099456-ae204035-bf14-400d-92ea-eb47e2ffd5ae.PNG)

![stocks_2018](https://user-images.githubusercontent.com/86337475/126099633-bb8246cb-6e71-40e0-a1de-26b526a204ac.PNG)

Using three output arrays for our ticker array, we were able to output total volume and calculate return by dividing endingPrice by startingPrice and subtracting that number by 1.

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

 ###**Execution Times**
 
 Initially, before refactoring, the VBA script took about 0.5s to run. After refactoring, it was down to around 0.1s. While not a lot, this is still a significant reduction in run time. 
 
 
 #**Summary**
 
 ###**Refactor Return on Time**
 The obvious advantage of refactoring code is that it will usually run faster. An important disadvantage to consider, however, is the time it will take to actually refactor. This applies to refactoring the original VBA script in that it took me much longer to write the code that would make the script run 0.4s faster. This is something that should be considered when refactoring code in general- how much time are you getting back by investing time in refactoring code to make it faster? In other workds, what is your refactor return for a particular project?
