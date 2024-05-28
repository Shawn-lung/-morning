import openpyxl
import yfinance as yf
from concurrent.futures import ThreadPoolExecutor
import subprocess
import os
import pandas as pd
import numpy as np


current_dir = os.path.dirname(os.path.abspath(__file__))

def open_and_close_excel(file_path):
    try:
        subprocess.Popen([file_path], shell=True)
        subprocess.call("timeout 5", shell=True)
    except Exception as e:
        print(f"Error: {e}")

def fetch_stock_data(stock_code):
    stock = yf.Ticker(stock_code)
    stock_info = stock.info
    trailingEps = stock_info.get('trailingEps', np.nan)
    beta = stock_info.get('beta', np.nan)
    forwardPE = stock_info.get('forwardPE', np.nan)
    returnOnEquity = stock_info.get('returnOnEquity', np.nan)
    returnOnAssets = stock_info.get('returnOnAssets', np.nan)
    grossMargin = stock_info.get('grossMargins', np.nan)
    operatingMargin = stock_info.get('operatingMargins', np.nan)
    peRatio = stock_info.get('trailingPE', np.nan)
    pbRatio = stock_info.get('priceToBook', np.nan)
    revenuePerShare = stock_info.get('revenuePerShare', np.nan)
    rtnList = [stock_code, trailingEps, beta, forwardPE, returnOnEquity, returnOnAssets, grossMargin, operatingMargin, peRatio, pbRatio, revenuePerShare]
    print(rtnList)
    return rtnList


file_path = os.path.join(current_dir, "taiex_mid100_stock_data.xlsx")

with open(os.path.join(current_dir, "tw_stock_codes.txt"), 'r') as file:
    lines = file.read().splitlines()
taiex_mid100_stocks = [line + ".TW" for line in lines]


with ThreadPoolExecutor() as executor:
    results = list(executor.map(fetch_stock_data, taiex_mid100_stocks))
try:
    workbook = openpyxl.load_workbook(file_path)
except FileNotFoundError:
    workbook = openpyxl.Workbook()


sheet_name = "Sheet"
if sheet_name not in workbook.sheetnames:
    sheet = workbook.create_sheet(sheet_name)
else:
    sheet = workbook[sheet_name]

sheet["A1"] = "股票代碼"
sheet["B1"] = "EPS"
sheet["C1"] = "Beta"
sheet["D1"] = "PE Ratio"
sheet["E1"] = "ROE"
sheet["F1"] = "ROA"
sheet["G1"] = "Gross margin"
sheet["H1"] = "Operating margin"
sheet["I1"] = "P/E Ratio"
sheet["J1"] = "P/B Ratio"
sheet["K1"] = "Revenue per share"

for i, result in enumerate(results, start=2):
    sheet[f"A{i}"] = result[0]
    sheet[f"B{i}"] = result[1]
    sheet[f"C{i}"] = result[2]
    sheet[f"D{i}"] = result[3]
    sheet[f"E{i}"] = result[4]
    sheet[f"F{i}"] = result[5]
    sheet[f"G{i}"] = result[6]
    sheet[f"H{i}"] = result[7]
    sheet[f"I{i}"] = result[8]
    sheet[f"J{i}"] = result[9]
    sheet[f"K{i}"] = result[10]

workbook.save(file_path)
open_and_close_excel(file_path)
