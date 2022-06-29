# Stock-analysis
## Project Overview
I was aksed by Steve to expand on the Green_Stocks workbook, that Steve loves. The expansion was done so Steve could do a litle more research for his parents. A macro was created using VBA to pull information from the "2017", and "2018" worksheets in the workbook into another worksheet named AllStocksAnalysis so Steve can run an analysis for all stocks. This worksheet will show Steve the ticker, total daily volume and the return for the chosen year. 
## Purpose
The purpose of this challenge was to refactor the VBA code used in the AllStockAnalysis and make it run more efficiently. In the future the dataset might grow and there will be data to parse. Refactoring the code the macro should run quicker. 
### Results
The code ran quicker after the refactoring. This was accomplished in part by creating 3 new output arrays and looping through those named arrays.
### The image below is the code run time on the "2017" dataset after refactoring.

![VBA Challenge Time 2017](https://github.com/pcar22/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

The code run time for the original code is 0.8359375
### The image below is the code run time on the "2018" dataset after refactoring.

![VBA Challenge Time 2018](https://github.com/pcar22/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The code run time for the original code is 0.9648438

## Summary
The advantage of refactoring a code is the time saving on a large dataset, there really isn't a disadvantage, however, it does take more time to solve the refactoring, and the time savings on a smaller dataset isn't really noticeable.
