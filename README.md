# VBA of Wall Street

## Overview of Project
For this week's project, we are using Visual Basic for Applications, also known as VBA, to analyze green energy stock data for the years of 2017 and 2018. VBA is a programming language that works with Excel so we can automate tasks and perform complex analyses. Instead of using Excel alone, we can now write scripts to create macros, or sub routines, that will allow us to read and write to cells, process functions, perform calculations, and generate tables and graphs. Throughout Module 2, we created a solution to compare 12 different green energy stocks, but we are looking to refactor the code to run faster. To make the code run more efficiently, we improve the logic of the existing code while keeping the same functionality.

### Purpose
The purpose of this project was to help Steve and his parents analyze green energy stock data. Although Steve’s parents already invested in Daqo New Energy Corporation, Steve wanted to analyze 11 other green energy stock options to compare their total daily volume and yearly return for 2017 and 2018. While we already had VBA code to read and write out our findings for the 12 green energy stocks, we also refactored the code to accommodate for a larger dataset with more stocks. Steve and his parents will now be able to view data for thousands of stocks and will be more informed on which green energy stock to invest in.

## Results
The analysis performed compares the total daily volumes and year returns for 12 green energy stocks. We initially used our VBA scripting to account for 12 stocks but decided to refactor our code to account for a larger dataset.

### Stock Analysis of Total Daily Volumes and Year Returns for 2017 and 2018
![VBA_Stock_Analysis_2017](/Resources/VBA_Stock_Analysis_2017.PNG) <img height="350" hspace="20"/> ![VBA_Stock_Analysis_2018](/Resources/VBA_Stock_Analysis_2018.PNG)

For this project, we wanted to look at the total daily volumes and year returns for the 12 green energy stocks above. The total daily volumes show often the stock was traded. This was calculated by summing up all the daily volumes for each stock. The yearly return shows the percentage increase or decrease in in stock price. We first had to find the price of the stock at the beginning of the year and the price of the stock at the end of the year. Then we divided the ending price by the starting price and subtracted 1 to get the yearly return for each stock. The two charts above show our results for 2017 and 2018. 

#### Findings for 2017
- All stocks, except for TERP, had a positive annual return
- Based on this chart alone, I don’t see a correlation between the total daily volumes and the yearly returns 
- DQ had the highest positive return of 199.4% but also the lowest amount of stocks traded

#### Findings for 2018
- All stocks, except ENPH and RUN, had a negative annual return
- ENPH and RUN both had high volume of stocks traded
- DQ dropped 62.6% but had more stocks traded in comparison to 2017

### Conclusion on Findings
Based on our findings for 2017 and 2018, I do not think we have enough information to make a conclusion on which green energy stocks to invest in. In 2017, it was a mostly positive year across the board for all the green energy stocks analyzed. In 2018, it was the opposite where most stocks had a negative annual return. We would need to do more research to see if there were any outside factors or events that may have occurred those years that would cause an overall affect on the stock market. Based on the total daily volumes, we cannot say there is a direct correlation between high volume of stocks traded and how well the stock is doing.

## Summary

### Advantages and Disadvantages of Refactoring Code in General

### Advantages and Disadvantages of the Original and Refactored VBA Script
