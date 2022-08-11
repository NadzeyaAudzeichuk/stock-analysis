# Stock Analysis with VBA

## Overview of Project

We have a workbook prepared beforehand as a starting point for this challenge. At a click of a button, it analyses a dozen stocks for their daily traded volumes and yearly returns for  2017 and 2018. Although the code works well for a dozen stocks, it may not work well for hundreds of stocks. And if it does, it may take a long time to execute. For this reason, we will go over previously written code and edit it in a way that will not matter if we are calculating ten or hundreds of stocks. Without adding new functionalities, we will restructure the internal logic of the code only.

### Purpose

This project aims to do CODE REFACTORING to determine whether it improves the logic of the code and makes the VBA script run-time more efficient. 

## Results

The data set contains 12 stocks with the range of prices and their volume traded on a particular date. Overall we have 3,012 rows. To calculate their daily volumes and yearly returns, we have **_Tickers()_** array and use _nested loops_ to loop through the data. We have a script that calculates the elapsed execution time to measure the code performance.

Within the _i-loop_, we go through **_Tickers()_** array and within the _j-loop_ through all of the data. We output the data for the current ticker before proceeding to the next one:
     ```
    'Loop through the Tickers() array
    For i = 0 To 11

        'Loop through all of the data
        For j = 2 To rowCount

        ...
            
        Next j

    'Output data for current ticker
    ...   
    
    Next i
    ```
The downside of the code here is that we loop over the whole data set (3,012 rows) for every ticker (12 times), which extends the processing time. Elapsed times to execute 2017 and 2018 data are 0.918 and 0.902 seconds, respectively:
![VBA_Challenge_2017.png](https://github.com/NadzeyaAudzeichuk/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)![VBA_ChallEnge_2018.png](https://github.com/NadzeyaAudzeichuk/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

To refactor the code, we create output arrays that will store results for the sum of the daily volumes, beginning, and end of the year ticker's prices:

    ```
    'Three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    ```  

We loop through the data one time and, using _if-then_ statements and _logical operators_, collect all the information. And save it in the arrays:

    ```
    'Loop to initialize the tickerVolumes
    For i = 0 To 11
      
        'tickerVolumes
         
    Next i
    
    'Loop over all the rows in the spreadsheet.
    For i = 2 To rowCount

        'Increase volume for current ticker
        'Get current tickerStartingPrice       
        'Get current tickerEndingPrice
        'Proceed to the next ticker
                
    Next i
    ```

 We output the content of the arrays in a spreadsheet:

    ```
    'Loop through out arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
      
        'tickers(i)
        'tickerVolumes(i)
        'tickerEndingPrice(i) / tickerStartingPrice(i) - 1

    Next i
    ```

With the script above, execution of 2017 and 2018 data takes 0.215 and 0.218 seconds, respectively:
![VBA_Challenge_2017_Refactored.png](https://github.com/NadzeyaAudzeichuk/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Refactored.png)![VBA_ChallEnge_2018_Refactored.png](https://github.com/NadzeyaAudzeichuk/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored.png)

As we can see, code refactoring improved the run time. The code above works more efficiently because it iterates the data one time only. Therefore, fewer steps are performed, and less memory is used to execute. 

## Summary

It is not easy to define what is good code and what is not. So refactoring has advantages and disadvantages. In general, by transforming the code, it executes more efficiently - takes fewer steps, uses less memory, helps find bugs, or improves the logic of the code to make it easier to read. In our example, the refactoring improved the run time. The script is executed four times faster, from ~0.8 to ~0.2 seconds.

On the other side, code refactoring could introduce new bugs or errors into the code; it might be time-consuming and risky. As for the data we have, the disadvantage is the complexity of multiple arrays. Because we have only 12 stocks and calculating two parameters, it is not hard to keep track of all calculations and arrays. In case of the need for more measures to be calculated, the complexity of the logic may increase. The original script would be easier to read and comprehend in this case.