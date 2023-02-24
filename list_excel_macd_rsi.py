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