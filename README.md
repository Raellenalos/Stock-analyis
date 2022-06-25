# Stock-analyis
## Purpose of the Analysis 
Steve recently graduated from college with his finance degree.
His parents are passionate about green energy. 
They have decided to invest in DAQO Energy Corp (DQ)  a company that makes silicon wafers for solar panels with little to no research. 
Steve wants to look in to DQ stocks for his parents
To diversify their funds, Steve wants to Analyze a handful of green energy stock in addition to DAQO stocks
Steve has created an excel file with the stock data showcasing 12 different stocks and their daily trade volume for the years of 2017 and 2018  
Steve wants to find the total daily volume and yearly return of each stock; the total number of shares traded throughout the day; percentage difference in price from the beginning of the year to the end of the year
Steve's parents want to know how actively DQ was traded in 2018. They believe that the more often a stock is traded than the price will accurately the value of the stock



## Analysis 

**Step 1a:**
Create a tickerIndex variable and set it equal to zero before iterating over all the rows. 

**Step 1b:**
Create three output arrays:
  * tickerVolumes as Long Data type 
  * tickerStartingPrices As Single 
  * tickerEndingPrices As Single 
  
**Step 2a:**
Create a for loop to initialize the tickerVolumes to zero.

**Step 2b:**
Create a for loop that will loop over all the rows in the spreadsheet.

**Step 3a:**
Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.

## Results:
### DQ 2017 Performance vs   DQ 2018 performance

In 2018 DQ went down 63%

![image](https://github.com/Raellenalos/Stock-analyis/blob/main/Resources/DQ%20Analysis%202017%20vs%202018.png)

* Original run of the data displaying the run time for 2017 and 2018

![image](https://github.com/Raellenalos/Stock-analyis/blob/main/Resources/All%20Stocks%202017.png)

![image](https://github.com/Raellenalos/Stock-analyis/blob/main/Resources/All%20stocks%202018.png)


## Summary: 

 **What are the advantages and disadvantage of refactoring code?**
 
 * Refactoring code can improve the code performance/ run time.
 
 **How do these Pros and cons apply to refactoring the original VBA script?**
 
* Code refactoring can be challenging becuase you have to keep in mind the order of the code and when you change or add a loop or any other function it has to be closed in the right order for the code to run.
