import os
import re
import sys
import time
import json
import urllib.request
from playwright.sync_api import sync_playwright
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass

KNOWN_STOCKS = {
    "SOXL", "MSTR", "NVDA", "TSLA", "AAPL", "AMD", "MSFT", 
    "AMZN", "GOOGL", "META", "COIN", "PLTR", "ARM", "SMCI",
    "NFLX", "DIS", "BA", "INTC", "QCOM", "SPY", "QQQ"
}

_BINGX_CONTRACTS_CACHE = None

def get_bingx_contracts_data():
    """
    快取並取得 BingX 全量 Swap 官方合約清單 (免金鑰 Public API)
    包含 1000+ 幣種之精確 symbol (如 NCSKSOXL2USD-USDT, 1000PEPE-USDT, BTC-USDT 等)
    """
    global _BINGX_CONTRACTS_CACHE
    if _BINGX_CONTRACTS_CACHE is not None:
        return _BINGX_CONTRACTS_CACHE
    
    url = "https://open-api.bingx.com/openApi/swap/v2/quote/contracts"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=4) as resp:
            res_json = json.loads(resp.read().decode("utf-8"))
            if res_json.get("code") == 0 and "data" in res_json:
                _BINGX_CONTRACTS_CACHE = res_json["data"]
                return _BINGX_CONTRACTS_CACHE
    except Exception as e:
        print(f"⚠️ [BingX] 獲取官方合約列表 API 提示: {e}，將自動改用內建啟發式規格比對。")
    
    _BINGX_CONTRACTS_CACHE = []
    return _BINGX_CONTRACTS_CACHE

def get_bingx_symbol_candidates(symbol_input: str, asset_type: str = None):
    """
    智能解析 BingX 候選網址 Symbol：
    1. 優先透過 BingX Swap 官方合約清單精準比對 (支援美股 NCSK、1000x 迷因幣、USDC 永續等)
    2. 備援採用啟發式與雙向 Fallback 規則 (保證斷網或新增未收錄幣種時依然可用)
    """
    raw = symbol_input.upper().strip()
    sym_clean = raw.replace("/", "").replace("_", "").replace("-", "")
    
    # 判斷基礎幣種與計價幣種
    if sym_clean.endswith("USDC"):
        base_unit = sym_clean[:-4]
        quote_unit = "USDC"
    elif sym_clean.endswith("USDT"):
        base_unit = sym_clean[:-4]
        quote_unit = "USDT"
    elif sym_clean.endswith("USD"):
        base_unit = sym_clean[:-3]
        quote_unit = "USDT"
    else:
        base_unit = sym_clean
        quote_unit = "USDT"

    standard_coin_key = f"{base_unit}USDT"
    
    # 若輸入本身已是 NCSK 格式 (例如 NCSKSOXL2USD-USDT)
    if raw.startswith("NCSK") and raw.endswith("-USDT"):
        match = re.search(r'NCSK(.*?)2USD-USDT', raw)
        base_unit = match.group(1) if match else raw.split("-")[0]
        standard_coin_key = f"{base_unit}USDT"
        return [raw], base_unit, standard_coin_key

    # 嘗試從 data/coin_types.json 讀取標的種類設定 (若存在)
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

    candidates = []
    contracts = get_bingx_contracts_data()

    # 1. 精準比對官方合約
    if contracts:
        # A. 完全符合 symbol (移除連字符後)
        for c in contracts:
            c_sym = c.get("symbol", "").upper().replace("-", "").replace("_", "")
            if sym_clean == c_sym:
                candidates.append(c["symbol"])
                break
        
        # B. 比對 displayName (例如 SOXL-USDT -> NCSKSOXL2USD-USDT)
        if not candidates:
            for c in contracts:
                c_disp = c.get("displayName", "").upper().replace("-", "").replace("_", "")
                if sym_clean == c_disp or f"{sym_clean}USDT" == c_disp:
                    candidates.append(c["symbol"])
                    break

        # C. 比對 1000/1000000 迷因幣格式 (例如 PEPE -> 1000PEPE-USDT)
        if not candidates:
            for c in contracts:
                c_sym = c.get("symbol", "").upper()
                if c_sym in [f"1000{base_unit}-{quote_unit}", f"1000000{base_unit}-{quote_unit}", f"10000{base_unit}-{quote_unit}"]:
                    candidates.append(c["symbol"])
                    break

        # D. 比對美股 NCSK 標的
        if not candidates:
            for c in contracts:
                if c.get("symbol", "").upper() == f"NCSK{base_unit}2USD-{quote_unit}":
                    candidates.append(c["symbol"])
                    break

    # 2. 啟發式規則備用擴展 (依標的類型排定優先順序)
    is_stock = (asset_type == 'stock') or (base_unit in KNOWN_STOCKS)
    if is_stock:
        stock_cand = f"NCSK{base_unit}2USD-{quote_unit}"
        if stock_cand not in candidates:
            candidates.append(stock_cand)
        std_cand = f"{base_unit}-{quote_unit}"
        if std_cand not in candidates:
            candidates.append(std_cand)
    else:
        std_cand = f"{base_unit}-{quote_unit}"
        if std_cand not in candidates:
            candidates.append(std_cand)
        # 迷因幣格式備援 (1000, 1000000)
        m1000 = f"1000{base_unit}-{quote_unit}"
        if m1000 not in candidates:
            candidates.append(m1000)
        # 美股 NCSK 備選
        stock_cand = f"NCSK{base_unit}2USD-{quote_unit}"
        if stock_cand not in candidates:
            candidates.append(stock_cand)

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
        max_lev = None
        for r in r_tiers:
            if m["min"] >= r["min"] and m["min"] < r["max"]:
                max_lev = int(r["lev"]) if r["lev"].is_integer() else r["lev"]
                break
        
        # 若超出現有檔位最高區間，依照最高檔位槓桿或 MMR 嚴格計算上限
        if max_lev is None:
            if r_tiers:
                base_lev = r_tiers[-1]["lev"]
                if m["mmr"] > 0:
                    mmr_cap = int(1.0 / m["mmr"])
                    max_lev = min(int(base_lev), mmr_cap)
                else:
                    max_lev = int(base_lev)
            elif m["mmr"] > 0:
                max_lev = int(1.0 / m["mmr"])
            else:
                max_lev = 100

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
    備援：透過 DOM 表格解析 BingX 數據 (支援多表格與固定表頭拆分結構)
    """
    rows_data = []
    tables = page.locator("table").all()
    for table_element in tables:
        tr_elements = table_element.locator("tbody tr, tr").all()
        for tr in tr_elements:
            tds = tr.locator("td").all()
            if len(tds) >= 4:
                texts = [td.inner_text().strip() for td in tds]
                # 若無數字則為表頭，自動跳過
                if not re.search(r'\d+', texts[0]):
                    continue
                tier_raw = texts[0]
                m = re.search(r'\d+', tier_raw)
                tier_str = f"檔位 {m.group()}" if m else tier_raw
                qty = clean_range_upper_bound(texts[1])
                mmr = texts[2]
                ded = texts[3]
                
                rows_data.append({
                    "交易所": "BingX",
                    "檔位": tier_str,
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
        browser = None
        for ch in ["chrome", "msedge"]:
            try:
                browser = p.chromium.launch(
                    channel=ch,
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
                )
                print(f"🌐 [BingX] 成功調用系統瀏覽器通道: {ch}")
                break
            except Exception:
                pass

        if not browser:
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
            )

        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1600, "height": 950},
            locale="zh-TW"
        )
        
        page = context.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")

        current_api_data = {}
        def handle_response(response):
            if "marginTiered" in response.url or "contract/marginTiered/get" in response.url:
                try:
                    res_json = response.json()
                    if res_json.get("code") == 0 and "data" in res_json:
                        current_api_data.update(res_json["data"])
                except Exception:
                    pass

        page.on("response", handle_response)
        
        for coin in coins:
            candidates, base_unit, standard_coin_key = get_bingx_symbol_candidates(coin)
            df = pd.DataFrame()

            for symbol_candidate in candidates:
                url = f"https://bingx.com/zh-tc/tradeInfo/perpetual/maintenance-margin-ratio/{symbol_candidate}/"
                print(f"\n--------------------------------------------------")
                print(f"🔄 [BingX] 連線幣種 [{standard_coin_key}] 候選網址: {url}")

                current_api_data.clear()

                try:
                    # 使用 domcontentloaded 快速加載，大幅縮短載入等待時間
                    page.goto(url, wait_until="domcontentloaded", timeout=12000)

                    # 輪詢等待 API 數據或 DOM 渲染完成 (最多 2.5 秒)
                    for _ in range(15):
                        if current_api_data:
                            break
                        page.wait_for_timeout(150)

                    if current_api_data:
                        df = parse_bingx_api_data(current_api_data, base_unit)
                        
                    if df.empty:
                        df = parse_bingx_dom_table(page)

                    if not df.empty:
                        print(f"✅ [BingX] 成功擷取 [{standard_coin_key}] (使用網址 {symbol_candidate}) 共 {len(df)} 檔數據！")
                        break
                    else:
                        print(f"⚠️ [BingX] [{symbol_candidate}] 未取得有效數據，嘗試下一個候選格式...")

                except Exception as e:
                    # 即使超時，若已截獲 API 數據則補救解析
                    if current_api_data:
                        df = parse_bingx_api_data(current_api_data, base_unit)
                        if not df.empty:
                            print(f"✅ [BingX] (超時補救成功) 擷取 [{standard_coin_key}] 共 {len(df)} 檔數據！")
                            break
                    print(f"❌ [BingX] 擷取 [{symbol_candidate}] 發生異常: {e}")

            # 雙重相容儲存 key，保證 main.py 能以任何格式正確獲取
            results[standard_coin_key] = df
            results[coin] = df
            results[coin.upper()] = df

            if not df.empty and excel_writer:
                sheet_name = standard_coin_key[:31]
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            elif df.empty:
                print(f"⚠️ [BingX] [{standard_coin_key}] 所有候選網址格式皆無有效數據。")

        browser.close()

    if excel_writer:
        excel_writer.close()
        print(f"\n🎉 BingX 數據抓取完成！檔案已儲存至 data/bingx_position_tier.xlsx")

    return results

if __name__ == "__main__":
    crawl_bingx_position_tiers()
