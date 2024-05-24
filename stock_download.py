import openpyxl
import yfinance as yf
from concurrent.futures import ThreadPoolExecutor
import subprocess
import os
import requests
import pandas as pd

def open_and_close_excel(file_path):
    try:
        subprocess.Popen([file_path], shell=True)
        subprocess.call("timeout 5", shell=True)
    except Exception as e:
        print(f"Error: {e}")


def fetch_stock_data(stock_code):
    stock = yf.Ticker(stock_code)
    stock_info = stock.info
    return stock_code, stock_info.get('trailingEps', 'NaN'), stock_info.get('beta', 'NaN'), stock_info.get('forwardPE', 'NaN'), stock_info.get('returnOnEquity', 'NaN'), stock_info.get('returnOnAssets', 'NaN')


file_path = os.path.join(os.getcwd(), "taiex_mid100_stock_data.xlsx")

with open(os.path.join(os.getcwd(), "tw_stock_codes.txt"), 'r') as file:
    lines= file.read().splitlines()
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

for i, result in enumerate(results, start=2):
    sheet[f"A{i}"] = result[0]
    sheet[f"B{i}"] = result[1]
    sheet[f"C{i}"] = result[2]
    sheet[f"D{i}"] = result[3]
    sheet[f"E{i}"] = result[4]
    sheet[f"F{i}"] = result[5]

workbook.save(file_path)
open_and_close_excel(file_path)