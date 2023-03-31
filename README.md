# YahooStocks
Guide to install python and python libraries:\n
https://medium.com/@carlosat/python-start-programming-today-an-introductory-guide-d1c750729dc2

You will need Python and the following libraries installed for this to work:
  * datetime
  * yfinance
  * mplfinance
  * pandas
  * xlsxwriter
  * openpyxl
  
Download data from Yahoo Finance using the yfinance library and create an excel workbook from the data.

The code will prompt the user for:
  1. The amount of stocks they want to browse.
  2. The tickers they wish to download data for.
  3. Start and end date for the tickers data.
  4. The period interval to download this data.
  
Code will then download the data with the set parameters to a single 'output.xlsx' file. 

Lastly it will prompt the user if they want to generate Candle Graphs for the stocks searched. If yes, it will create a .png image '(stock) price graph.png'
