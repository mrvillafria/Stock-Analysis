# VBA of Wall Street

## Overview of Project
For this week's project, we are using Visual Basic for Applications, also known as VBA, to analyze green energy stock data for the years of 2017 and 2018. VBA is a programming language that works with Excel so we can automate tasks and perform complex analyses. Instead of using Excel alone, we can now write scripts to create macros, or sub routines, that will allow us to read and write to cells, process functions, perform calculations, and generate tables and graphs. Throughout Module 2, we created a solution to compare 12 different green energy stocks, but we are looking to refactor the code to run faster. To make the code run more efficiently, we improved the logic of the existing code while keeping the same functionality.

### Purpose
The purpose of this project was to help Steve and his parents analyze green energy stock data. Although Steve’s parents already invested in Daqo New Energy Corporation, Steve wanted to analyze 11 other green energy stock options to compare their total daily volume and yearly return for 2017 and 2018. While we already had VBA code to read and write out our findings for the 12 green energy stocks, we also refactored the code to accommodate for a larger dataset with more stocks. Steve and his parents will now be able to view data for thousands of stocks and will be more informed on which green energy stock to invest in.

## Results
The analysis performed compares the total daily volumes and year returns for 12 green energy stocks. We initially used our VBA scripting to account for 12 stocks but decided to refactor our code to account for a larger dataset.

### Stock Analysis of Total Daily Volumes and Year Returns for 2017 and 2018
For this project, we wanted to look at the total daily volumes and year returns for the 12 green energy stocks pictured below. The total daily volumes show how often the stock was traded. This was calculated by summing up all the daily volumes for each stock. The yearly return shows the percentage increase or decrease in the stock price. We first had to find the price of the stock at the beginning of the year and the price of the stock at the end of the year. Then we divided the ending price by the starting price and subtracted 1 to get the yearly return for each stock. If there was a positive annual return, we changed the cell to be green and if there was a negative annual return, we changed the cell to be red. The two charts below show our results for 2017 and 2018. 

|2017     |2018      |
|------------|-------------|
| ![VBA_Stock_Analysis_2017](/Resources/VBA_Stock_Analysis_2017.PNG)| ![VBA_Stock_Analysis_2018](/Resources/VBA_Stock_Analysis_2018.PNG)|

#### Findings for 2017
- All stocks, except for TERP, had a positive annual return
- Based on this chart alone, I don’t see a correlation between the total daily volumes and the yearly returns 
- DQ had the highest positive return of 199.4% but also the lowest amount of stocks traded

#### Findings for 2018
- All stocks, except ENPH and RUN, had a negative annual return
- ENPH and RUN both had high volume of stocks traded
- DQ dropped 62.6% but had more stocks traded in comparison to 2017

#### Conclusion on Stock Analysis
Based on our findings for 2017 and 2018, I do not think we have enough information to make a conclusion on which green energy stocks to invest in. In 2017, it was a mostly positive year across the board for all the green energy stocks analyzed. In 2018, it was the opposite where most stocks had a negative annual return. We would need to do more research to see if there were any outside factors or events that may have occurred those years that would cause an overall affect on the stock market. We could also run our script against more data if it was available to compare additional years. Based on the total daily volumes, we cannot say there is a direct correlation between high volume of stocks traded and how well the stock is doing.

### Analysis of Execution Times of the Original Script vs the Refactored Script
Throughout Module 2, we created a script to compare the 12 green energy stocks that Steven selected. While the original script worked for this data set, we refactored the script to loop through the data one time to collect all the information. We used VBA's Timer function to measure the code performance and a message box to display the elapsed time. 

|Original Code Execution Time     |Refactored Code Execution Time      |
|------------|-------------|
|![VBA_Challenge_Original_2017](/Resources/VBA_Challege_Original_2017.PNG)|![VBA_Challenge_2017](/Resources//VBA_Challenge_2017.PNG)|
|![VBA_Challenge_Original_2018](/Resources/VBA_Challenge_Original_2018.PNG)|![VBA_Challenge_2018](/Resources//VBA_Challenge_2018.PNG)|

#### Conclusion on Execution Times
Based on our findings, the refactored code saved time in comparision to the original code. This will help save processing time if Steven and his parents want to run a larger data set with more stocks.

## Summary

### Advantages and Disadvantages of Refactoring Code in General

#### Advantages
There are many advantages to refactoring code. As you can see from our analysis above, refactoring code to run more efficiently can save time when executing. This will be especially important if the code has a large dataset to run through. In the business world, time is valuable so people will be more satisfied if they can quickly receive their results. Another advantage is if you improve the logic of the code, it will be easier for someone else to look at it and troubleshoot in the future. Refactoring code is also beneficial because while you are reviewing the original code, you can find any bugs that need to be fixed.

#### Disadvantages
Although there are a lot of advantages to refactoring your code, one major downfall is the extra time it takes to edit the code. Not only do you have to think of better logic but there will always be additional troubleshooting. Depending on how long your deadline is to have the refactored code in place, it may be very risky.

### Advantages and Disadvantages of the Original and Refactored VBA Script
There are some clear advantages and disadvantages of the original script versus the refactored script. But first, I would like to outline the steps for both the original and refactored VBA scripts:

#### The original script went through the following steps:
1. Format the output sheet on the "All Stocks Analysis" worksheet.
2. Initialize an array of all tickers.
3. Prepare for the analysis of tickers.
    1. Initialize variables for the starting price and ending price.
    2. Activate the data worksheet.
    3. Find the number of rows to loop over. 
4. Loop through the tickers.
5. Loop through rows in the data.
    1. Find the total volume for the current ticker.
    2. Find the starting price for the current ticker.
    3. Find the ending price for the current ticker.
6. Output the data for the current ticker.

#### The refactored script went through the following steps:
1. Format the output sheet on All Stocks Analysis worksheet.
2. Initialize array of all tickers.
    1. Activate data worksheet.
    2. Get the number of rows to loop over.
3. Create a ticker Index.
4. Create 3 output arrays: tickerVolumes, tickerStartingPrices, tickerEndingPrices.
5. Create a for loop to initialize the tickerVolumes to zero.
6. Loop over all the rows in the spreadsheet.
    1. Increase volume for current ticker.
    2. Find the starting price by checking if the current row is the first row with the selected tickerIndex.
    3. Find the ending price by checking if the current row is the last row with the selected tickerIndex. 
    4. Increase the tickerIndex.
7. Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
8. Formatting

#### Advantages
While it may look like there are more steps for the refactored code, it runs faster than the original script. One big step that helped reduce the time was that the refactored scripting did not have a nested loop. The refactored script also has comments explaining what each step does that will help describe the process to the next person who looks at the script.

#### Disadvantages
One disadvantage of refactoring the code was editing and troubleshooting the script was very time consuming. Changing the logic of the script by adding 3 more arrays and removing the nested loop took longer for me to think through than expected.
