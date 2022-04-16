# VBA Challenge: Refactoring Stock Analysis Code
## Overview of Project: 

### Purpose
The initial analysis of the Stock data was targeted to analyze a dozen stocks, but it might not be an efficient code to run for a larger dataset. To expand the analysis, the code will need to be refactored to create a loop that loops over the data set once and collects all the information in one-go. This is a worthy task because if a code is written better then it would take less steps to execute the results, use less memory, and improve the readability. There are many ways to accomplish a task, but refactoring allows for the written code to be the most efficient way of getting results. 

### Background
This analysis was initially built because a recent finance graduate, Steve, needs to analyze if his first client's stock portfolio is sufficiently diversified. He hopes to be able to present the information in an easy-to-understand way to illustrate his findings. Based on the initial analysis of the one stock that his first client's invested in, he finds that the stock, DAQO dropped over 63% in a single year. Steve will need to offer better stocks to his clients, so he will need a complete analysis of multiple stocks within multiple years to find some options. The initial All Stocks Analysis wasn't designed to be flexible, so refactoring the code will allow for specific edits to make it useable now and in the future. 

## Results: 

**Changes to the Code:**

1. Creating a ticker index: `tickerindex = 0`
Creating this variable and setting it equal to zero before iterating over the dataset allowed for the correct index to be access across the arrays that will be introduced to the code. This action sets it to a specific position within the array; it does not set it equal to value 0. 


2. Creating arrays: 

    `Dim tickerVolumes(12) As Long; 
    Dim tickerStartingPrices(12) As Single; 
    Dim tickerEndingPrices(12) As Single` 
    
Three output arrays are created to store the data that matches set qualifications that will be outlined in the code. This will allow for the code to run smoother because as the correct data is found, it will be stored in the array, so the data will not have to be looped through multiple times to extract the pertinent information. The number (12) is present after each array name because it shows that each array can hold 12 values, which matches the 12 stocks being analyzed. With this code, we make an assumption that the data is in order and grouped together to allow for the following code to extract data correctly. The data types are specified as Long vs. Single because the `tickerVolumes` is going to be a very large value, while the Single data type is used because we don't need the full data width of Double. 
    
3. For Loops: The syntax of our For loops is important because we have to first initialize the `tickerVolumes` to zero using a For loop. 

[Initialize tickerVolumes to Zero]![image](https://user-images.githubusercontent.com/102566199/163690955-7ffdab17-021e-451b-8167-2fab8b3e1d59.png). 

Then a For loop that loops over all the rows must be written that allows for the increase in the variable `tickerVolumes` as it reaches the end of a ticker's data.

[For Loop to run through data and extract values]![image](https://user-images.githubusercontent.com/102566199/163691052-265ad536-294f-4750-81f4-f00b0309db82.png). 

Once the data is extracted from the sheet, it will be placed in a specified cell by assigning the array to that cell. 

4. If-Then Statements: The use of the If-Then statements allow the code to find the data within the set and store the data to the proper array. This is where it becomes important that the data is in order. If the data is out of order, the If-Then statements won't be able to assign the correct `tickerEndingPrices` to the proper tickers. The code, ` If Cells(i - 1, 1).Value <> tickers(tickerIndex)` and ` If Cells(i + 1, 1).Value <> tickers(tickerIndex)`, checks to make sure that the rows above and below the ticker information are different. This is a special way to make sure that the `tickerStartingPrice` and the `TickerEndingPrice` are asssigned at the very beginning of the ticker data and at the very end. 

