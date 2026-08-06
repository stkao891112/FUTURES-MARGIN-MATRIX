import os
import re
import sys
import time
import concurrent.futures
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from binance import crawl_binance_leverage_margin
from bybit import crawl_bybit_multi_coins_to_excel
from bitget import crawl_bitget_position_tiers
from okx import crawl_okx_position_tiers
from mexc import crawl_mexc_position_tiers

def extract_tier_number(val):
    """
    從檔位字串提取數字做為排序依據（例如 "檔位 1" -> 1, "2" -> 2）
    """
    if pd.isna(val):
        return 999999
    val_str = str(val).strip()
    match = re.search(r'\d+', val_str)
    return int(match.group()) if match else 999999

def clean_range_upper_bound(val_str):
    """
    清理範圍字串，僅保留上限數值（例如 "0~200,000" -> "200,000", "200,000~1,000,000" -> "1,000,000"）
    """
    if not isinstance(val_str, str):
        return val_str
    val_str = str(val_str).strip()
    for sep in ['~', '～']:
        if sep in val_str:
            parts = val_str.split(sep)
            if len(parts) > 1 and parts[-1].strip():
                return parts[-1].strip()
    return val_str

def normalize_dataframe_columns(df):
    """
    統一四大交易所的欄位名稱，以 Binance / 7大標準結構為基準對齊：
    1. 交易所
    2. 檔位
    3. 單位 (OKX視搜尋幣種，其餘固定USDT)
    4. 倉位分級
    5. 最大槓桿
    6. 維持保證金率
    7. 維持金額(USDT)
    """
    if df.empty:
        return df

    column_mapping = {}
    for col in df.columns:
        c_str = str(col).strip()
        if c_str == "交易所":
            column_mapping[col] = "交易所"
        elif c_str in ["檔位", "等級", "欄位_1"]:
            column_mapping[col] = "檔位"
        elif c_str == "單位":
            column_mapping[col] = "單位"
        elif any(k in c_str for k in ["倉位", "風險限額", "價值", "張數", "名義"]) or c_str == "欄位_2":
            column_mapping[col] = "倉位分級"
        elif "槓桿" in c_str or c_str == "欄位_3":
            column_mapping[col] = "最大槓桿"
        elif any(k in c_str for k in ["保證金"]) or c_str == "欄位_4":
            column_mapping[col] = "維持保證金率"
        elif any(k in c_str for k in ["金額", "扣減額"]) or c_str == "欄位_5":
            column_mapping[col] = "維持金額(USDT)"

    renamed_df = df.rename(columns=column_mapping)

    # 若缺乏「單位」欄位，預設補上 USDT
    if "單位" not in renamed_df.columns:
        renamed_df["單位"] = "USDT"

    # 針對倉位價值/範圍欄位自動萃取上限數值 (如 0~200,000 -> 200,000)
    if "倉位分級" in renamed_df.columns:
        renamed_df["倉位分級"] = renamed_df["倉位分級"].apply(clean_range_upper_bound)

    standard_headers = ["交易所", "檔位", "單位", "倉位分級", "最大槓桿", "維持保證金率", "維持金額(USDT)"]
    final_cols = [c for c in standard_headers if c in renamed_df.columns]
    extra_cols = [c for c in renamed_df.columns if c not in final_cols]
    return renamed_df[final_cols + extra_cols]

def merge_and_sort_exchanges_data(coin_dfs):
    """
    合併多個交易所的 DataFrame，將欄位統一對齊，
    並依據「交易所」名稱與「檔位」進行排序。
    """
    if not coin_dfs:
        return pd.DataFrame()

    normalized_dfs = [normalize_dataframe_columns(df) for df in coin_dfs if not df.empty]
    if not normalized_dfs:
        return pd.DataFrame()

    combined_df = pd.concat(normalized_dfs, ignore_index=True)

    # 確保「交易所」欄位在第 1 欄 (A 欄)
    if "交易所" in combined_df.columns:
        cols = ["交易所"] + [c for c in combined_df.columns if c != "交易所"]
        combined_df = combined_df[cols]

    # 依「交易所」與「檔位」排序
    if "檔位" in combined_df.columns and "交易所" in combined_df.columns:
        combined_df["__tier_sort"] = combined_df["檔位"].apply(extract_tier_number)
        combined_df.sort_values(by=["交易所", "__tier_sort"], ascending=[True, True], inplace=True)
        combined_df.drop(columns=["__tier_sort"], inplace=True)
    elif "交易所" in combined_df.columns:
        combined_df.sort_values(by=["交易所"], ascending=[True], inplace=True)

    return combined_df

def apply_excel_styles_and_colors(ws):
    """
    為 Excel 工作表套用色彩標籤與專業樣式：
    - Binance: 柔和黃色 (#FFF2CC)
    - Bitget: 柔水藍色 (#E0F2FE)
    - Bybit: 柔和橘色 (#FFE0B2)
    - OKX: 柔和淡紫 / 灰藍 (#E8EAF6)
    - 標題列: 深灰藍色 (#2C3E50) 與白色粗體
    """
    fill_binance = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # 幣安黃
    fill_bitget = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")   # Bitget 水藍
    fill_bybit = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")    # Bybit 橘
    fill_okx = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")      # OKX 淡紫
    fill_mexc = PatternFill(start_color="E0F2F1", end_color="E0F2F1", fill_type="solid")     # MEXC 青綠
    fill_default = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    fill_header = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="微軟正黑體", size=10)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    # 1. 樣式化標題列 (Row 1)
    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center
        cell.border = thin_border

    # 2. 逐行設定資料儲存格背景色與邊框
    for row in range(2, ws.max_row + 1):
        ex_val = str(ws.cell(row=row, column=1).value or "").strip()

        if "Binance" in ex_val:
            row_fill = fill_binance
        elif "Bitget" in ex_val:
            row_fill = fill_bitget
        elif "Bybit" in ex_val:
            row_fill = fill_bybit
        elif "OKX" in ex_val:
            row_fill = fill_okx
        elif "MEXC" in ex_val:
            row_fill = fill_mexc
        else:
            row_fill = fill_default

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = row_fill
            cell.font = font_data
            cell.border = thin_border
            
            if col in [1, 2, 3]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # 3. 自動調整欄寬
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            display_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
            max_len = max(max_len, display_len)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

def run_single_exchange(args):
    """
    線程池執行的單一交易所抓取封裝函式
    """
    name, crawl_func, coins = args
    print(f"\n⚡ [並行任務啟動] 開始抓取 [{name}] 數據...")
    start_time = time.time()
    try:
        res = crawl_func(coins, save_excel=False)
        elapsed = time.time() - start_time
        return name, res or {}, f"✅ 成功 (耗時 {elapsed:.1f} 秒)"
    except Exception as e:
        elapsed = time.time() - start_time
        return name, {}, f"❌ 失敗: {e} (耗時 {elapsed:.1f} 秒)"

def main():
    print("=" * 60)
    print(" 🚀 跨交易所合約保證金數據自動化爬蟲（Binance/Bybit/Bitget/OKX/MEXC 5家並行版） ")
    print("=" * 60)
    
    # 支援命令列參數傳入、coins.json 設定檔讀取或互動式輸入
    if len(sys.argv) > 1:
        coins = [c.upper() for c in sys.argv[1:] if c.strip()]
    elif os.path.exists(os.path.join("data", "coins.json")):
        try:
            with open(os.path.join("data", "coins.json"), "r", encoding="utf-8") as f:
                coins = json.load(f)
        except Exception:
            coins = ["BTCUSDT", "ETHUSDT", "SOLUSDT"]
    else:
        coins = ["BTCUSDT", "ETHUSDT", "SOLUSDT"]

    print(f"\n📌 本次任務將為以下 {len(coins)} 個幣種抓取五大交易所數據: {coins}")
    
    tasks = [
        ("Binance (幣安)", crawl_binance_leverage_margin, coins),
        ("Bybit", crawl_bybit_multi_coins_to_excel, coins),
        ("Bitget", crawl_bitget_position_tiers, coins),
        ("OKX", crawl_okx_position_tiers, coins),
        ("MEXC", crawl_mexc_position_tiers, coins),
    ]

    all_results = {}
    summary_results = {}

    print("\n" + "=" * 60)
    print(f" ⚡ 正在同時 (並行) 啟動 {len(tasks)} 家交易所 Playwright 瀏覽器...")
    print("=" * 60)

    # 使用 ThreadPoolExecutor 同時啟動 4 家爬蟲
    start_total_time = time.time()
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(tasks)) as executor:
        futures = [executor.submit(run_single_exchange, task) for task in tasks]
        for future in concurrent.futures.as_completed(futures):
            ex_name, res, status = future.result()
            all_results[ex_name] = res
            summary_results[ex_name] = status

    total_elapsed = time.time() - start_total_time

    # 進行同一幣種跨交易所數據整合
    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    combined_excel_filename = os.path.join(data_dir, "all_exchanges_futures_margin.xlsx")
    print("\n" + "=" * 60)
    print(" 🔀 正在進行同幣種跨交易所數據合併與排序...")
    print("=" * 60)

    with pd.ExcelWriter(combined_excel_filename, engine='openpyxl') as writer:
        for coin in coins:
            coin_dfs = []
            for ex_name, res_dict in all_results.items():
                if coin in res_dict and not res_dict[coin].empty:
                    coin_dfs.append(res_dict[coin])

            if coin_dfs:
                merged_df = merge_and_sort_exchanges_data(coin_dfs)
                sheet_name = coin[:31]
                merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 套用背景色彩與專業格式
                ws = writer.sheets[sheet_name]
                apply_excel_styles_and_colors(ws)
                print(f" 💾 已成功將 [{coin}] 4家交易所數據寫入工作表: [{sheet_name}] (共 {len(merged_df)} 列)")
            else:
                print(f" ⚠️ [{coin}] 無有效數據可供合併。")

    # 印出執行總結報告
    print("\n" + "=" * 60)
    print(" 📊 總結報告 (Task Summary)")
    print("=" * 60)
    for ex_name, status in summary_results.items():
        print(f" • {ex_name:<20}: {status}")
    print(f"\n⏱️  並行爬蟲總耗時: {total_elapsed:.1f} 秒")

    if os.path.exists(combined_excel_filename):
        size_kb = os.path.getsize(combined_excel_filename) / 1024
        print(f"\n🎉 整合完成！同幣種跨交易所匯總 Excel 已儲存至:")
        print(f" 📂 {os.path.abspath(combined_excel_filename)} (大小: {size_kb:.1f} KB)")
        
        # 自動轉出 JSON 供網頁版計算機實時使用
        try:
            from export_json import export_tiers_to_json
            export_tiers_to_json()
        except Exception:
            pass

if __name__ == "__main__":
    main()
