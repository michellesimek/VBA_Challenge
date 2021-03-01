# VBA_Challenge
Homework assignment #2 using VBA scripting within Excel to analyze real stock market data.

GoogleDrive Link to Excel Spreadsheet: https://drive.google.com/drive/folders/1pqB5qcuVxp7nzONJSo3QAtQMXIGlAB4f?usp=sharing

## VBA Script
This project included creating a VBA script that would run on every worksheet in the workbook by running the VBA script once.

The script loops through each worksheets and outputs the following:
* Ticker symbol
* Yearly change in price from opening to closing for each ticker
* Percent change in price from opening to closing for each ticker
* Total stock volume for each ticker 

### Ticker Symbol
To find the ticker symbol, a script was created to loop through row each in the first column. When the row below did not match the row above, the ticker symbol for the row above would appear in a running list in Column H.

### Yearly Change 
To find the yearly change in price from opening to closing, the opening price (listed as inital_price in script) was subtracted from the closing price.

This column was formatted to show positive changes in green and negative changes in red. 

### Percent Change
To find the percent change in price, the yearly change value was divided by the opening price (listed as inital_price in script) and turned into a percentage. 

### Total Stock Volume
The total stock volume of each ticker was found by adding the volume of each indiviudal date from each ticker.

### Summary Table of Results
A summary table was created to show which ticker has the greatest percent increase, greatest percent decrease, and greatest total stock volume.