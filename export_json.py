import os
import re
import json
import sys
from datetime import datetime
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def parse_tier_num(val):
    if pd.isna(val):
        return 1
    match = re.search(r'\d+', str(val))
    return int(match.group()) if match else 1

def parse_number(val):
    if pd.isna(val) or val == '-':
        return 0.0
    val_str = str(val).replace(',', '').replace('%', '').replace('x', '').strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def parse_mmr(val):
    if pd.isna(val) or val == '-':
        return 0.0
    val_str = str(val).strip()
    is_percent = '%' in val_str
    num_str = val_str.replace(',', '').replace('%', '').replace('x', '').strip()
    try:
        num = float(num_str)
        if is_percent or num > 1.0:
            return num / 100.0
        return num
    except ValueError:
        return 0.0

def export_tiers_to_json():
    excel_path = os.path.join("data", "all_exchanges_futures_margin.xlsx")
    if not os.path.exists(excel_path):
        print("⚠️ 未檢測到 [data/all_exchanges_futures_margin.xlsx] 匯總總表！")
        print("🚀 正在自動執行 main.py 爬取 5 大交易所最新合約檔位數據...")
        try:
            from main import main as run_main
            run_main()
        except Exception as e:
            print(f"❌ 執行 main.py 提示: {e}")

    if not os.path.exists(excel_path):
        print("⚠️ 無法讀取 Excel 檔，取消轉出 JSON。")
        return

    try:
        excel_file = pd.ExcelFile(excel_path)
        all_data = {}
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        all_data["_last_updated"] = timestamp_str

        for sheet in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            coin = sheet.upper()
            all_data[coin] = {
                "Binance": [],
                "Bitget": [],
                "Bybit": [],
                "OKX": [],
                "MEXC": [],
                "Hyperliquid": []
            }

            if coin == "BTCUSDT":
                all_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.0125, "deduction": 0, "maxLev": 50}]
            elif coin == "ETHUSDT":
                all_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.02, "deduction": 0, "maxLev": 50}]
            elif coin == "SOLUSDT":
                all_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.025, "deduction": 0, "maxLev": 25}]

            for _, row in df.iterrows():
                ex = str(row.get("交易所", "")).strip()
                if "Binance" in ex:
                    ex_key = "Binance"
                elif "Bitget" in ex:
                    ex_key = "Bitget"
                elif "Bybit" in ex:
                    ex_key = "Bybit"
                elif "OKX" in ex:
                    ex_key = "OKX"
                elif "MEXC" in ex:
                    ex_key = "MEXC"
                else:
                    continue

                tier = parse_tier_num(row.get("檔位", 1))
                limit_val = parse_number(row.get("倉位分級", 0))
                max_lev = parse_number(row.get("最大槓桿", 100))
                mmr_val = parse_mmr(row.get("維持保證金率", 0))
                ded_val = parse_number(row.get("維持金額(USDT)", 0))

                item = {
                    "tier": tier,
                    "mmr": mmr_val,
                    "deduction": ded_val,
                    "maxLev": max_lev
                }
                if ex_key in ["OKX", "MEXC"]:
                    item["limitQty"] = limit_val
                else:
                    item["limit"] = limit_val

                all_data[coin][ex_key].append(item)

        # 寫入 JSON
        json_path = os.path.join("data", "tiers.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)
        
        last_updated_path = os.path.join("data", "last_updated.json")
        with open(last_updated_path, "w", encoding="utf-8") as f:
            json.dump({"last_updated": timestamp_str}, f, ensure_ascii=False, indent=2)

        print(f"✅ 已成功將匯總總表檔位數據轉出至 JSON (包含時間戳 {timestamp_str})")

        # 同步嵌入更新 index.html 中的 let EXCHANGE_TIERS 與 let LAST_UPDATED
        html_path = "index.html"
        if os.path.exists(html_path):
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()
            
            tiers_json_str = json.dumps(all_data, ensure_ascii=False, indent=2)
            new_tiers_block = f"let EXCHANGE_TIERS = {tiers_json_str};"
            new_updated_block = f'let LAST_UPDATED = "{timestamp_str}";'
            
            updated_html = re.sub(
                r'let EXCHANGE_TIERS = \{.*?\};',
                new_tiers_block,
                html_content,
                flags=re.DOTALL
            )
            if 'let LAST_UPDATED =' in updated_html:
                updated_html = re.sub(
                    r'let LAST_UPDATED = ".*?";',
                    new_updated_block,
                    updated_html
                )

            with open(html_path, "w", encoding="utf-8") as f:
                f.write(updated_html)
            print(f"✅ 已成功將 6 大交易所全檔位數據與時間戳直嵌同步至 index.html！")

    except Exception as e:
        print(f"⚠️ 導出過程提示: {e}")

if __name__ == "__main__":
    export_tiers_to_json()
