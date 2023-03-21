import datetime
import yfinance as yf
import mplfinance as mpf
import pandas as pd
import xlsxwriter
import openpyxl
#creating our excel file
excelname = 'output.xlsx'
workbook = xlsxwriter.Workbook(excelname)
worksheet = workbook.add_worksheet()
#prompting the user for information on amount of stocks, ticker names, date ranges and interval.
while True:
    try:
        n = int(input('How Many stocks do you want to browse?: '))
        while n > 10:
            print('Thats too many stocks. You can search up to 10.')
            n = int(input('How Many stocks do you want to browse?: '))
    except ValueError:
        print('That input is not valid, try again.')
        continue
    else:
        break

stocks = []
for i in range(n):
    s = str(input(f'Input the ticker for stock {i+1}: '))
    stocks.append(s)

while True:
    try:
        start_date = datetime.date.fromisoformat(
            input(str('State your start date (YYYY-MM-DD): ')))
    except ValueError:
        print('That date is not valid, try again.')
        continue
    else:
        break

while True:
    try:
        end_date = datetime.date.fromisoformat(
            input(str('State your end date (YYYY-MM-DD): ')))
        end_date += datetime.timedelta(days=1) #we extend the end date by 1day because yfinance will exclude it otherwise
    except ValueError:
        print('That date is not valid, try again.')
        continue
    else:
        break

intervals = ('1d', '5d', '1mo', '3mo', '6mo', '1y', '2y', '5y', '10y', 'ytd', 'max')
while True:
    try:
        print('Valid Intervals are: 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max')
        interval = input(str('State your time interval: '))
        while (interval not in intervals) or (interval == ''):
            print('That is not a valid interval. Valid intervals are:\nDays: 1d 5d\nMonths: 1mo 3mo 6mo\nYears: 1y 2y 5y 10y ytd max')
            interval = input(str('State your time interval: '))
    except ValueError:
        print('That interval is not valid, try again.')
        continue
    else:
        break
#downloading our stock tickers data with yfinance and creating a pandas dataframe from it
data = yf.download(stocks, start=start_date, end=end_date, group_by='ticker',interval=interval).tz_localize(None)
pdata = pd.DataFrame(data)
pstats = pdata.describe() #getting basic description of our data with pandas describe() function
with pd.ExcelWriter(excelname) as writer: #sending our prices and statistics sheet to the excel file
    data.to_excel(writer,sheet_name = 'prices')
    pstats.to_excel(writer,sheet_name='statistics')

y = ['y','Y','yes','Yes','YEs','YES','YeS','yES','yEs','yeS']
r = input('Do you want to generate Candle Graphs of your stocks? Y/N: ') #prompting the user if graphs should be created
for i in range(n):
    t = yf.Ticker(stocks[i])
    shr = t.get_shares_full(start=start_date,end=end_date)
    if shr is not None: #some stocks dont have shares or dividend data, if thats the case the data is not manipulated
        shares = pd.DataFrame(t.get_shares_full(start=start_date,end=end_date).tz_localize(None))
    dvd = t.dividends        
    if dvd is not None:
        dividends = pd.DataFrame(dvd.tz_localize(None))
    #writing to the excel file shares and dividend data
    with pd.ExcelWriter(excelname,engine='openpyxl',mode='a') as writer:
        if shr is not None:
            shares.to_excel(writer,sheet_name = f'{stocks[i]} shares')
        if dvd is not None:
            dividends.to_excel(writer,sheet_name = f'{stocks[i]} dividends')
    #creating price graphs if user had requested it before
    if r in y:
        pgraph = yf.download(stocks[i],start=start_date, end=end_date,interval=interval).tz_localize(None)
        plot = mpf.plot(pgraph,type='candle',title= f'{stocks[i]}',style='yahoo',savefig=f'{stocks[i]} price graph.png')
