# -*- coding: utf-8 -*-
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import json
import os
import requests
import pandas as pd
from datetime import datetime
import warnings

CONFIG_PATH = 'config.json'

def load_config():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def download_file(name, url, save_path):
    try:
        print(f"[{name}] 下載中：{url}")
        response = requests.get(url)
        response.raise_for_status()
        with open(save_path, 'wb') as f:
            f.write(response.content)
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{name}] 下載完成 → {save_path}")
        return True
    except Exception as e:
        print(f"[{name}] 下載失敗：{e}")
        return False

def convert_excel_to_csv(input_file, output_folder):
    try:
        warnings.simplefilter("ignore", UserWarning)

        if not os.path.exists(input_file):
            print(f"[ERROR] 檔案不存在：{input_file}")
            return

        file_ext = os.path.splitext(input_file)[-1].lower()
        engine = 'xlrd' if file_ext == '.xls' else 'openpyxl'

        xls = pd.ExcelFile(input_file, engine=engine)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"[INFO] 建立資料夾：{output_folder}")

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, engine=engine, dtype=str)
            safe_sheet_name = "".join([c if c.isalnum() else "_" for c in sheet_name])
            csv_path = os.path.join(output_folder, f"{safe_sheet_name}.csv")
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"[SUCCESS] 工作表 '{sheet_name}' → {csv_path}")

        print("[INFO] 所有工作表轉換完成。")

    except Exception as e:
        print(f"[WARNING] 轉檔失敗：{e}")

def run_once():
    config = load_config()
    for key, info in config.items():
        # 若還保留 Download_Time 欄位就跳過
        if key == "Download_Time":
            continue
        url = info.get("Path")
        save_name = info.get("Save_name")
        if url and save_name:
            success = download_file(key, url, save_name)
            if success:
                output_folder = f"csv_output/{key}"
                convert_excel_to_csv(save_name, output_folder)

if __name__ == "__main__":
    run_once()
