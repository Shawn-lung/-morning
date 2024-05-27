import yfinance as yf
import pandas as pd


stock = yf.Ticker('2330.TW')
print(stock.history(period = '1y'))