

import pandas, xlsxwriter # xlsxwriter to write data into excel
import yfinance as yf   # yfinance to get historical data



stock_symbols = ['ACC', 'ADANIPORTS', 'ADANIENT', 'ADANIPOWER', 'AMBUJACEM', 'APOLLOHOSP', 'ARVIND', 'ASIANPAINT',
'AUROPHARMA', 'BAJFINANCE', 'BALKRISIND', 'BANKBARODA', 'BANKINDIA', 'BERGEPAINT', 'BHEL', 'BAJAJFINSV',
'BOSCHLTD', 'CADILAHC', 'CENTURYTEX', 'CHOLAFIN', 'CUMMINSIND', 'DIVISLAB', 'DLF', 'ENGINERSIN', 'EQUITAS',
'ESCORTS', 'FEDERALBNK', 'GODREJCP', 'GRASIM', 'HAVELLS', 'HCLTECH', 'BSOFT', 'HDFCBANK', 'HINDPETRO',
'CANBK', 'HINDZINC', 'IDEA', 'IDFCFIRSTB', 'CASTROLIND', 'INDIGO', 'INDUSINDBK', 'INFY', 'IOC', 'DHFL',
'ITC', 'JINDALSTEL', 'EICHERMOT', 'JSWSTEEL', 'JUBLFOOD', 'JUSTDIAL', 'KOTAKBANK', 'EXIDEIND', 'L&TFH',
'LUPIN', 'M&M', 'GLENMARK', 'MARICO', 'MCDOWELL-N', 'MCX', 'MOTHERSUMI', 'MRF', 'MUTHOOTFIN',
'NATIONALUM', 'NCC', 'ICICIPRULI', 'NESTLEIND', 'NTPC', 'ONGC', 'PAGEIND', 'INDUSTOWER', 'PFC', 'PIDILITIND',
'PNB', 'RAMCOCEM', 'RAYMOND', 'RECLTD', 'RELCAPITAL', 'RELIANCE', 'SAIL', 'SIEMENS', 'SRF', 'SRTRANSFIN',
'STAR', 'TATACONSUM', 'TATAPOWER', 'TCS', 'MGL', 'TECHM', 'MINDTREE', 'TORNTPOWER', 'UBL', 'UJJIVAN',
'UNIONBANK', 'UPL', 'VEDL', 'VOLTAS', 'WIPRO', 'YESBANK', 'AMARAJABAT', 'COFORGE', 'RELINFRA', 'APOLLOTYRE',
'AXISBANK', 'TITAN', 'ASHOKLEY', 'BAJAJ-AUTO', 'BATAINDIA', 'BEL', 'BHARATFORG', 'BIOCON',
'BPCL', 'BRITANNIA', 'CESC', 'CIPLA', 'COALINDIA', 'DABUR', 'DISHTV', 'GAIL', 'GMRINFRA', 'HDFC', 'HEROMOTOCO',
'CONCOR', 'HINDALCO', 'DRREDDY', 'HINDUNILVR', 'IBULHSGFIN', 'IDBI', 'IGL', 'MANAPPURAM', 'MFSL',
'M&MFIN', 'MARUTI', 'OFSS', 'NMDC', 'OIL', 'POWERGRID', 'RBLBANK', 'SHREECEM', 'SUNPHARMA',
'TATASTEEL', 'TATAELXSI', 'TORNTPHARM', 'TVSMOTOR', 'ULTRACEMCO', 'ZEEL', 'COLPAL', 'ICICIBANK',
'LICHSGFIN', 'SBIN', 'BHARTIARTL', 'LT', 'NBCC', 'KAJARIACER', 'PEL', 'PETRONET', 'PVR', 'TATACHEM',
'TATAMOTORS', 'TATAMTRDVR', 'SUNTV']

# Reanamed 
# NIITTECH to COFORGE
# TATAGLOBAL TO TATACONSUM
# INFRATEL TO INDUSTOWER

# Removed 'HEXAWARE' share 
# Because HEXAWARE Share is delisted on November 3, 2020 so the stock price for last 90 days is not available 



append_str = '.NS'
Stocks = [stock + append_str for stock in stock_symbols] # Appending '.NS' to all symbols to get Indian Stock price
Stocks




writer = pd.ExcelWriter('Stock Info.xlsx', engine='xlsxwriter')  #  creating a new excel file named Stock info.xlsx
 # pd.ExcelWriter for writing DataFrame objects into excel sheets.

for stock in Stocks:  # iterating through Stocks list
    
    data = yf.download(stock , start='2021-06-04', end='2021-10-13')['Close'] 
    # passing stock symbol as tickers argument and specifying start & end date for downloading Historical Stock price data
    df = 0
    df = pd.DataFrame(data, columns={'Close'})
    df['Close'] = df['Close'].round(decimals = 2)  # rounding of Close to 2 decimal places
    df.to_excel(writer, sheet_name=stock)  # passing the stock name as sheet_name
    worksheet = writer.sheets[stock]
    worksheet.set_column('A1:A1',20)  # setting column width of Date column to size 20
    
writer.save()  # saving the Excel file

# 'Stock Info.xlsx' file has 160 sheets 
# Each sheet(Stock) has its Closing price for last 90 days 





