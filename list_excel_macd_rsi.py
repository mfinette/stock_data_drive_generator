import yfinance as yf
import pandas as pd
import ta
import json
import requests

# Read the tickers from list.txt file
with open('list.txt') as f:
    tickers = [line.strip().upper() for line in f]

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter('stocks_data.xlsx', engine='xlsxwriter')

for ticker in tickers:
    sheet_name = ticker

    # Get the data from Yahoo Finance
    if ticker == 'NDAQ' or ticker == '^GSPC':
        data = yf.download(ticker, period='max')
    else:
        data = yf.download(ticker, period='4y')

    # Calculate MACD
    exp1 = data['Close'].ewm(span=12, adjust=False).mean()
    exp2 = data['Close'].ewm(span=26, adjust=False).mean()
    data['MACD'] = exp1 - exp2
    data['MACD_signal'] = data['MACD'].ewm(span=9, adjust=False).mean()

    # Calculate RSI
    delta = data['Close'].diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()
    rs = avg_gain / avg_loss
    data['RSI'] = 100 - (100 / (1 + rs))

    # Calculate RSI weekly
    data['RSI_weekly'] = data['RSI'].rolling(window=7).mean()

    # Calculate weekly and monthly MACD
    data['MACD_weekly'] = data['Close'].ewm(span=12, adjust=False).mean() - data['Close'].ewm(span=26*5, adjust=False).mean()
    data['MACD_monthly'] = data['Close'].ewm(span=12, adjust=False).mean() - data['Close'].ewm(span=26*22, adjust=False).mean()

    # Calculate weekly and monthly MACD signals
    data['MACD_signal_monthly'] = data['MACD_monthly'].ewm(span=9, adjust=False).mean()
    data['MACD_signal_weekly'] = data['MACD_weekly'].ewm(span=9, adjust=False).mean()

    # Calculate stochastic oscillator
    data['%K_daily_5'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=5, smooth_window=3)
    data['%D_daily_5'] = data['%K_daily_5'].rolling(window=3).mean()
    data['%K_weekly_5'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=5*7, smooth_window=3)
    data['%D_weekly_5'] = data['%K_weekly_5'].rolling(window=3).mean()
    data['%K_monthly_5'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=5*30, smooth_window=3)
    data['%D_monthly_5'] = data['%K_monthly_5'].rolling(window=3).mean()

    data['%K_daily_89'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=89, smooth_window=3)
    data['%D_daily_89'] = data['%K_daily_89'].rolling(window=3).mean()
    data['%K_weekly_89'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=89*7, smooth_window=3)
    data['%D_weekly_89'] = data['%K_weekly_89'].rolling(window=3).mean()
    # Format date column
    data['date'] = data.index.date.astype(str)

    # Reorder columns
    data = data[['date', 'Open', 'High', 'Low', 'Close', 'Volume', 'MACD', 'MACD_signal', 'MACD_weekly', 'MACD_signal_weekly', 'MACD_monthly', 'MACD_signal_monthly', '%K_daily_5', '%D_daily_5', '%K_weekly_5', '%D_weekly_5', '%K_monthly_5', '%D_monthly_5','%K_daily_89', '%D_daily_89', '%K_weekly_89', '%D_weekly_89', 'RSI', 'RSI_weekly']]
    data = data.sort_values(by='date', ascending=False)

    # Write to Excel
    data.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer.close()

# Upload to drive
var = input("Send file? (yes or no)")
if var.lower() == 'yes':
    headers = {
        "Authorization":"Bearer ya29.a0AVvZVsrn_7TMng1MC67opmomm7DRZmcjIDT9tIegvEnYPd6i5g94ZA6hzwuRVPD1P4qLCZFnPmBZoMRuUx52mqKfijG7rRaJYOfhv6fgHmCRHKpcAkry-bkXe9OxaPEJnectEWq-N989S8RmZOGACflcJ9JOaCgYKATISARASFQGbdwaIhvfYqFojy1UT8uH4oCGjzA0163"
    }
    
    para = {
        "name":"stocks_data.xlsx",
        "parents":["1ABeELo6Abu84K-ymnVDZZQt2dFhkHTJp"]
    }
    
    files = {
        'data':('metadata',json.dumps(para),'application/json;charset=UTF-8'),
        'file':open('./stocks_data.xlsx','rb')
    }
    
    r = requests.post("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers=headers,
        files=files
    )
    print(r)