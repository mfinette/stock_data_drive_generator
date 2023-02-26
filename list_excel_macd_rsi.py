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

    # Calculate stochastic oscillators
    data['%K'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=14, smooth_window=3)
    data['%D'] = data['%K'].rolling(window=3).mean()
    data['%K_89'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=89, smooth_window=3)
    data['%D_89'] = data['%K_89'].rolling(window=3).mean()
    data['%K_5'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=5, smooth_window=3)
    data['%D_5'] = data['%K_5'].rolling(window=3).mean()

    # Format date column
    data['date'] = data.index.date.astype(str)

    # Reorder columns
    data = data[['date', 'High', 'Low', 'Close', 'Volume', 'MACD', 'MACD_signal', 'MACD_weekly', 'MACD_signal_weekly', 'MACD_monthly', 'MACD_signal_monthly', '%K', '%D', '%K_89', '%D_89', '%K_5', '%D_5', 'RSI']]
    data = data.sort_values(by='date', ascending=False)

    # Write to Excel
    data.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer.close()

# Upload to drive
var = input("Send file? (yes or no)")
if var.lower() == 'yes':
    headers = {
        "Authorization":"Bearer ya29.a0AVvZVsqJT0lATEV_CawvtbD5pqiNkNgQRDUhYJmKjiqoJej_rO5yYarwMpGVmAwDMToSxaRKvzZ1Ors0840RP9OjDx39A1vMmMp3rjtaxoPvHUZWlUvJuti-b7MyTY_f-tZaYNV6dsYplv6HCM-3FwEiNAW1aCgYKAVsSARASFQGbdwaIrYq0WPtQBq8fJNppCBwZIA0163"
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