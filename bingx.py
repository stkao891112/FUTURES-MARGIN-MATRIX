import os
import re
import sys
import time
import json
from playwright.sync_api import sync_playwright
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def get_bingx_symbol_candidates(symbol_input: str, asset_type: str = None):
    """
    將輸入的幣種字串（如 BTC, BTCUSDT, BTC-USDT, SOXLUSDT）轉換為 BingX 可能的網址 Symbol 候選清單：
    - 標準加密貨幣格式: BTC-USDT
    - 美股/指數合約格式: NCSKSOXL2USD-USDT
    """
    sym = symbol_input.upper().strip()
    
    if sym.startswith("NCSK") and sym.endswith("-USDT"):
        match = re.search(r'NCSK(.*?)2USD-USDT', sym)
        base_unit = match.group(1) if match else sym.split("-")[0]
        standard_coin_key = f"{base_unit}USDT"
        return [sym], base_unit, standard_coin_key

    if sym.endswith("_USDT"):
        sym = sym.replace("_USDT", "-USDT")
    elif sym.endswith("USDT") and "-" not in sym:
        base = sym[:-4]
        sym = f"{base}-USDT"
    elif "-" not in sym:
        sym = f"{sym}-USDT"

    base_unit = sym.split("-")[0]
    standard_coin_key = f"{base_unit}USDT"

    # 若未直接指定 asset_type，嘗試從 data/coin_types.json 讀取標的種類設定
    if not asset_type:
        try:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            types_path = os.path.join(BASE_DIR, "data", "coin_types.json")
            if os.path.exists(types_path):
                with open(types_path, "r", encoding="utf-8") as f:
                    types_data = json.load(f)
                    asset_type = types_data.get(standard_coin_key) or types_data.get(base_unit)
        except Exception:
            pass

    if asset_type == 'stock':
        candidates = [
            f"NCSK{base_unit}2USD-USDT",
            f"{base_unit}-USDT"
        ]
    else:
        candidates = [
            f"{base_unit}-USDT",
            f"NCSK{base_unit}2USD-USDT"
        ]
    return candidates, base_unit, standard_coin_key

def clean_range_upper_bound(val_str: str) -> str:
    """
    清理範圍字串，僅保留上限數值（例如 "0 ~ 320,000" -> "320,000"）
    """
    if not isinstance(val_str, str):
        return val_str
    val_str = str(val_str).strip()
    for sep in ['~', '～', '-']:
        if sep in val_str:
            parts = val_str.split(sep)
            if len(parts) > 1 and parts[-1].strip():
                return parts[-1].strip()
    return val_str

def parse_bingx_api_data(api_data: dict, base_unit: str):
    """
    解析 BingX 離線 API 數據 (newMaintenanceTiered / maintenanceTiered + marginRiskLevelTiered)
    """
    m_str = api_data.get("newMaintenanceTiered") or api_data.get("maintenanceTiered", "")
    r_str = api_data.get("marginRiskLevelTiered", "")
    
    if not m_str:
        return pd.DataFrame()

    m_tiers = []
    for item in m_str.split(";"):
        if not item.strip():
            continue
        parts = item.split(":")
        if len(parts) >= 2:
            range_part = parts[0]
            mmr_part = parts[1]
            ded_part = parts[3] if len(parts) > 3 else "0"
            if "-" in range_part:
                min_v, max_v = range_part.split("-")
                m_tiers.append({
                    "min": float(min_v),
                    "max": float(max_v),
                    "mmr": float(mmr_part),
                    "deduction": float(ded_part)
                })

    r_tiers = []
    if r_str:
        for item in r_str.split(";"):
            if not item.strip():
                continue
            parts = item.split(":")
            if len(parts) >= 2 and "-" in parts[0]:
                range_part = parts[0]
                lev_part = parts[1]
                min_v, max_v = range_part.split("-")
                r_tiers.append({
                    "min": float(min_v),
                    "max": float(max_v),
                    "lev": float(lev_part)
                })

    rows_data = []
    for idx, m in enumerate(m_tiers):
        max_lev = 100
        for r in r_tiers:
            if m["min"] >= r["min"] and m["min"] < r["max"]:
                max_lev = int(r["lev"]) if r["lev"].is_integer() else r["lev"]
                break
        
        limit_str = f"{int(m['max']):,}" if m['max'].is_integer() else f"{m['max']:,}"
        mmr_str = f"{m['mmr'] * 100:.2f}%".rstrip('0').rstrip('.') if m['mmr'] * 100 % 1 != 0 else f"{int(m['mmr'] * 100)}%"
        if m['mmr'] * 100 < 1:
            mmr_str = f"{m['mmr'] * 100:.4f}%".rstrip('0').rstrip('.')
        
        ded_str = f"{int(m['deduction']):,}" if m['deduction'].is_integer() else f"{m['deduction']:,}"

        rows_data.append({
            "交易所": "BingX",
            "檔位": f"檔位 {idx + 1}",
            "單位": "USDT",
            "倉位分級": limit_str,
            "最大槓桿": f"{max_lev}x",
            "維持保證金率": mmr_str,
            "維持金額(USDT)": ded_str
        })

    return pd.DataFrame(rows_data)

def parse_bingx_dom_table(page):
    """
    備援：透過 DOM 表格解析 BingX 數據
    """
    rows_data = []
    table_element = page.locator("table").first
    if table_element.count() > 0:
        tr_elements = table_element.locator("tbody tr, tr").all()
        for tr in tr_elements:
            tds = tr.locator("td").all()
            if len(tds) >= 4:
                texts = [td.inner_text().strip() for td in tds]
                if "檔位" in texts[0] or "Tier" in texts[0]:
                    continue
                tier = texts[0]
                qty = clean_range_upper_bound(texts[1])
                mmr = texts[2]
                ded = texts[3]
                
                rows_data.append({
                    "交易所": "BingX",
                    "檔位": tier,
                    "單位": "USDT",
                    "倉位分級": qty,
                    "最大槓桿": "-",
                    "維持保證金率": mmr,
                    "維持金額(USDT)": ded
                })
        if rows_data:
            return pd.DataFrame(rows_data)
    return pd.DataFrame()

def crawl_bingx_position_tiers(coins=None, save_excel=True):
    if coins is None:
        raw_input = input("請輸入想要查詢 BingX 檔位的幣種 (用空白分隔，例如: BTCUSDT ETHUSDT SOXLUSDT): ").strip().upper()
        if not raw_input:
            print("❌ 幣種名稱不可空白！程式結束。")
            return {}
        coins = raw_input.split()
    elif isinstance(coins, str):
        coins = coins.strip().upper().split()
    elif isinstance(coins, list):
        coins = [c.strip().upper() for c in coins if c.strip()]

    if not coins:
        print("❌ 幣種名稱不可空白！程式結束。")
        return {}

    print(f"\n📋 [BingX] 即將依次查詢以下 {len(coins)} 個幣種: {coins}")

    results = {}
    excel_writer = None
    if save_excel:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        data_dir = os.path.join(BASE_DIR, "data")
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        excel_filename = os.path.join(data_dir, "bingx_position_tier.xlsx")
        excel_writer = pd.ExcelWriter(excel_filename, engine='openpyxl')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        for coin in coins:
            candidates, base_unit, standard_coin_key = get_bingx_symbol_candidates(coin)
            df = pd.DataFrame()

            for symbol_candidate in candidates:
                url = f"https://bingx.com/zh-tc/tradeInfo/perpetual/maintenance-margin-ratio/{symbol_candidate}/"
                print(f"\n--------------------------------------------------")
                print(f"🔄 正在連線至 BingX 幣種 [{symbol_candidate}]: {url}")

                page = context.new_page()
                api_data = {}

                def handle_response(response):
                    nonlocal api_data
                    if "contract/marginTiered/get" in response.url or "marginTiered" in response.url:
                        try:
                            res_json = response.json()
                            if res_json.get("code") == 0 and "data" in res_json:
                                api_data = res_json["data"]
                        except Exception:
                            pass

                page.on("response", handle_response)
                
                try:
                    page.goto(url, wait_until="networkidle", timeout=25000)
                    page.wait_for_timeout(1500)

                    if api_data:
                        df = parse_bingx_api_data(api_data, base_unit)
                        
                    if df.empty:
                        df = parse_bingx_dom_table(page)

                    if not df.empty:
                        print(f"✅ 成功擷取 [{standard_coin_key}] (使用網址 {symbol_candidate}) 共 {len(df)} 檔 BingX 數據！")
                        page.close()
                        break
                    else:
                        print(f"⚠️ [{symbol_candidate}] 未取得數據，嘗試下一個候選網址格式...")

                except Exception as e:
                    print(f"❌ 擷取 [{symbol_candidate}] 發生異常: {e}")
                finally:
                    if not page.is_closed():
                        page.close()

            results[standard_coin_key] = df

            if not df.empty and excel_writer:
                sheet_name = standard_coin_key[:31]
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            elif df.empty:
                print(f"⚠️ [{standard_coin_key}] 所有人候選網址格式皆無有效 BingX 數據。")

        browser.close()

    if excel_writer:
        excel_writer.close()
        print(f"\n🎉 BingX 數據抓取完成！檔案已儲存至 data/bingx_position_tier.xlsx")

    return results

if __name__ == "__main__":
    crawl_bingx_position_tiers()
