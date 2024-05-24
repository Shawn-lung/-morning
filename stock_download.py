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
# 定義臺灣中型100指數的成分股列表
taiex_mid100_stocks = [
    "2330.TW", "2317.TW", "2454.TW", "1301.TW", "2882.TW", "2881.TW", 
    "3008.TW", "2303.TW", "5880.TW", "2886.TW", "2891.TW", "2885.TW", 
    "2880.TW", "2002.TW", "1303.TW", "2892.TW", "2883.TW", "2884.TW", 
    "6505.TW", "9910.TW", "3711.TW", "2357.TW", "1216.TW", "2412.TW", 
    "2603.TW", "1326.TW", "2308.TW", "2890.TW", "1402.TW", "2615.TW", 
    "2912.TW", "2327.TW", "2207.TW", "2353.TW", "2409.TW", "2609.TW", 
    "6415.TW", "3697.TW", "2382.TW", "1101.TW", "3034.TW", "3231.TW", 
    "2231.TW", "6180.TW", "1590.TW", "1605.TW", "3474.TW", "4904.TW", 
    "2379.TW", "8046.TW", "3037.TW", "2347.TW", "3293.TW", "2498.TW", 
    "6116.TW", "8454.TW", "6285.TW", "2345.TW", "3532.TW", "5269.TW", 
    "6269.TW", "1536.TW", "4763.TW", "1783.TW", "6239.TW", "2823.TW", 
    "1476.TW", "4938.TW", "3682.TW", "1722.TW", "6456.TW", "6206.TW", 
    "5904.TW", "2227.TW", "3661.TW", "3596.TW", "6176.TW", "6271.TW", 
    "4105.TW", "2362.TW", "6488.TW", "1760.TW", "4919.TW", "4551.TW", 
    "8299.TW", "3406.TW", "6121.TW", "3017.TW", "2633.TW", "6282.TW", 
    "4736.TW", "4743.TW"
]


file_path = os.path.join(os.getcwd(), "taiex_mid100_stock_data.xlsx")

res = requests.get("https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=1&issuetype=1&industry_code=&Page=1&chklike=Y")
df = pd.read_html(res.text)[0][2]
df.to_csv(os.path.join(os.getcwd(),"tw_stock_codes.csv"), index=False)



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