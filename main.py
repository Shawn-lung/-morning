import subprocess
import os
from stock_download import StockDownloader
from score import StockScorer

def main():
    stock_list_file = "tw_stock_codes.txt"
    input_file = "taiex_mid100_stock_data.xlsx"

    downloader = StockDownloader(stock_list_file, input_file)
    downloader.process()

    scorer = StockScorer(input_file)
    scorer.process()
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, input_file)

    try:
        subprocess.Popen([file_path], shell=True)
        subprocess.call("timeout 5", shell=True)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

