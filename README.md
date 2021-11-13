# Analyze Stock Data with VBA

### Project Overview

Steve is a friend whose parents are his first financial clients.  
Steve’s parents are invested in a Green Energy company called DAQO Energy Corp.  
Steve would like to analyze select stock data in the same sector as DAQO Energy Corp to better assist his parents with increasing their return on investment (ROI).  
The purpose of this project is to automate the stock data analysis using VBA.

### DAQO Analysis
Steve provided a spreadsheet containing multiple years of stock data for the 12 stocks in the Green Energy sector in which his parents are interested in investment.  
I used the formula Closing Price/Opening Price -1 to calculate the return on investment (ROI).  
The snapshots of the DAQO data shows a manual search of the DAQO data.  
Manually searching the spreadsheet for DAQO opening price and closing price is inefficient.  

![This is a alt text.](/image/sample.png "This is a sample image.")

### Examine VBA

A for loop was utilized to iterate through the rows for stock data searching for the DQ ticker from 2018 stock data.  
The nested if statements find the initial opening price and final closing price of DAQO, store then to variables, then calculated the annual return on investment.

![This is a alt text.](/image/sample.png "This is a sample image.")

### DQ Analysis Results

The initial analysis of DAQO Energy Corp showed a 63% loss for return on investment.  
The results can be viewed below.  This significant loss was not an acceptable ROI for Steve’s parents.

### All Stocks Analysis

I created a VBA scrip to analyze all stock data provided by Steve.  
AllStocksAnalysis prompts the user for the year then returns the ROI for all the stock data and runtime. Below you can see the results for 2017 and 2018.

![This is a alt text.](/image/sample.png "This is a sample image.")

### Refactoring Code

I was able to refactor the initial DQ stock analysis code to create the All-Stock analysis code in VBA.  

Refactoring the code has some benefits:
* I did not have to start over, I could expand the functionality of the previously used subroutines and looks
* It saved time to add formatting and features to existing code
* It makes the code more comprehensible to any other developer who inspects it
* I was able to decrease the runtime despite adding formatting and user controls such as buttons
Refactoring code has multiple advantages but one major disadvantage, time is spend modifying code that is functional to improve readability and use in the future.  Most developers under a time constraint may not want to spend time refactoring code that is already functional to solve future problems that do not currently exist. 

### All Stocks Analysis Results

Below you can see the results of running the refactored VBA code for 2017 and 2018.  The runtime has decreased by 80%.  The formatting allows the user to quickly determine which stocks had a positive return on investment as they are highlighted in green.  Most of the stocks in the Green Energy Sector that Steve selected performed better in 2017 than 2018.  After reviewing the Stock analysis data in the refactored VBA code, I would recommend that Steve’s parents consider ENPH and RUN as green energy company investment alternatives to DQ.  

![This is a alt text.](/image/sample.png "This is a sample image.")







