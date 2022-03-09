# Module 2 Deliverable 2: Refactoring VBA Code

## Overview of Project

Steve had previously requested that I make a workbook for his parents so they may analyze some green stocks from 2017 and 2018. The previous workbook had included the total daily volume of the stock sales as well as their average yearly returns at the click of a button. Steve loved the workbook and now wants to be able to do research on the entire stock market over the last few years.

### Purpose

The purpose of this challange is to refactor the prior code to loop through all the data one time in order to collect the same information, namely the total daily volume of the stock sales and their average yearly returns.

## Analysis and Challenges

For this refactored code, I created the tickerIndex variable and set it equal to zero before iterating over all the rows. Using the same tickerIndex, I accessed the correct index across the four different arrays I used: the tickers array and the three output arrays which are tickerVolumes, tickerStartingPrices and tickerEndingPrices. The first 'for' loop was to initialize the tickerVolumes to zero and subsequently the loop will loop over all the rows in the spreadsheet. Inside the previous 'for' loop I wrote a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current volume for the current stock ticker. I used the tickerIndex variable as the index. Using row numbering with the selected tickerIndex, I later used if-then statements to assign tickerStartingPrices and tickerEndingPrices. Using a 'for' loop through the 4 arrayes, I gave the output assignments as "Ticker", "Total Daily Volume" and "Return" columns on the spreadsheet.
<img width="1110" alt="green_stocks_2017_Timer" src="https://user-images.githubusercontent.com/95712234/157379467-53c2c73e-59ef-485d-90a6-2d7f5cf6968e.png">
<img width="1069" alt="green_stocks_2018_Timer" src="https://user-images.githubusercontent.com/95712234/157379487-6a512f47-f74b-4fbd-936c-b7d8b271550a.png">

Improved timings with the refactored code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95712234/157379503-faa07b97-22fc-4121-a318-610edd8d976c.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95712234/157379515-6e7c5ba6-5ab1-4ade-9026-993bbcbfcdc3.png)


## Results

- The workbook is now capable of analyzing larger sets of data at any given year more efficiently.

- Based on the results of the timer for the previous green_stock workbook and the current VBA_Challenge workbook, we can see that the refactored code is much more time efficient in its analysis.



## Summary

Based on my experience with this project, I have observed that there are both advantages and disadvantages of refactoring code in general. The main advantage of refactoring would be that it gives the data analyst a template to work with. The disadvantage can be seen in new challanges, such as having to recreate a button as it stopped working.

Personally, I found working with an original script more rewarding as it gave a 'clean slate' for creativity. Refactoring seemed to constrict the process and at times, frustrating when trying to maneauver the code or sticking to refactoring instructions, such as variable names I would have otherwise not chosen (tickerVolume vs. totalVolume).
