# VBA-Challenge
Summary
The code in this .bas file is used as a macro in excel to run an analysis on stock market data. It generates a summary table to the right of the data on each sheet of the workbook and compiles the tickers in the data set with their yearly change, percent change from open to close in that year, and the total stock volume of all the transactions in that year. The yearly change is color-coded to help identify when it was positive (green) and negative (red). 

Limitations
The code in this .bas file depends on the data to be formatted in a certain way for this code to run as expected. The spreadsheet must include headers with column A including the tickers, column B including the date, column c including the open price, column d including the high, column e including the low, column f including the close price, and column g including the volume. The high (column d) and low (column e) are not used in this analysis. 

The data must be formatted with tickers **a to z** or **z to a** and the dates with oldest to newest for each ticker. The code relies on the change between the tickers and pulls the first open date assuming that it is the first date within the year and the last closing date of that ticker assuming it is the last date of the year. 

Methodology
1. Sets a for loop to run through each worksheet within the workbook using.
1. Establishes the important variables for both the formatting of the summary table (called tickerSummary) as well as the calculations and values that must be set through the module/macro.
1. Establishes the start value before the for loop to capture the correct value.
1. Start a for loop to run through the tickers, open price and closing price and add up the volume.
1. If the ticker has changed then the code sets the ticker name and the closing price. 
1. If the start value is not zero then percent change between the opening and closing price is calculated, if it is zero it is set as such in the summary.
1. It sets the total volume number adding previous volume numbers for when the ticker had not changed.
1. It reports out the ticker and yearly change values
1. It color codes yearly change using an if statement.
1. It reports out the percent change and adjusts the format of that column to a percent. 
1. It reports out the total volume for that ticker.
1. It resets the start to the next ticker's opening value.

Screenshots are uploaded as files
![Screen shot 2014](https://github.com/andersonmalloryk/VBA-Challenge/blob/main/Anderson_VBA_2014.jpg)
![Screen shot 2015](https://github.com/andersonmalloryk/VBA-Challenge/blob/main/Anderson_VBA_2015.jpg)
![Screen shot 2016](https://github.com/andersonmalloryk/VBA-Challenge/blob/main/Anderson_VBA_2016.jpg)
