import yfinance as yf
import pandas as pd
import ta
import json
import requests

########################################################################################################################################################
# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter('stocks_data.xlsx', engine='xlsxwriter')

tickers = []
while True:
    ticker = input("Enter a ticker symbol (enter 'done' when finished): ")
    if ticker.lower() == 'done':
        break
    tickers.append(ticker.upper())

for ticker in tickers:
    sheet_name = ticker

    # Get the data from Yahoo Finance
    data = yf.download(ticker, period='max')

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

    # Calculate weekly and monthly MACD
    data['MACD_weekly'] = data['Close'].ewm(span=12, adjust=False).mean() - data['Close'].ewm(span=26*5, adjust=False).mean()
    data['MACD_monthly'] = data['Close'].ewm(span=12, adjust=False).mean() - data['Close'].ewm(span=26*22, adjust=False).mean()

    # Calculate stochastic oscillator
    data['%K'] = ta.momentum.stoch(data['High'], data['Low'], data['Close'], window=14, smooth_window=3)
    data['%D'] = data['%K'].rolling(window=3).mean()

    # Format date column
    data['date'] = data.index.date.astype(str)

    # Reorder columns
    data = data[['date', 'High', 'Low', 'Close', 'Volume', 'MACD', 'MACD_signal', 'MACD_weekly', 'MACD_monthly', '%K', '%D', 'RSI']]
    data = data.sort_values(by='date', ascending=False)

    # Write to Excel
    data.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer.close()
########################################################################################################################################################

#drive part :DDDDDDDDDDDDDDDDDD
var = input("Send file? (yes or no)")
if var.lower() == 'yes':
    headers = {
        "Authorization":"Bearer ya29.a0AVvZVsqXb-ltLAtCXI5CdzSsoGQt3Vqig1hmFtYbMccVh3JoGWM6ILLZA3omQPedAZV5XNBw9id-5a3cyM9SktIj6Iqv9w-ADgCQB0rHJ9IzmU8utp3wV-RxWC1wcfIQuUYReFKPN7TwN6qE7SeVpG0sN9q_aCgYKAcoSARASFQGbdwaIYtvJsU2y_sacbDX5ZKG6vg0163"
    }
    
    para = {
        "name":"stocks_data.xlsx",
        "parents":["1uXwXat5YKMD71koxv8Ox6dPRkUT9hGun"]
    }
    
    files = {
        'data':('metadata',json.dumps(para),'application/json;charset=UTF-8'),
        'file':open('./stocks_data.xlsx','rb')
    }
    
    r = requests.post("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers=headers,
        files=files
    )