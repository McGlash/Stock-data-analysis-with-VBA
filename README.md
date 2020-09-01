# The VBA of Wall Street

## Background

I used VBA scripting to analyze real stock market data. 

### Data

* [Stock Data](Resources/Multiple_year_stock_data.xlsx)

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Approach

* Created a scripts that looped through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Used conditional formatting to highlight positive change in green and negative change in red.

* The result looked as follows.

![moderate_solution](Images/moderate_solution.png)

* Returned the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution looked as follows:

![hard_solution](Images/hard_solution.png)
