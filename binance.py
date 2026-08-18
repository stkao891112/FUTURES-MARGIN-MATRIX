import os
import re
import sys
from playwright.sync_api import sync_playwright
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def clean_range_value(val_str: str) -> str:
    """
    清理範圍字串，例如將 "0 - 300,000" 自動轉換為 "300,000"。
    """
    if not isinstance(val_str, str):
        return val_str
    
    if '-' in val_str:
        match = re.search(r'-\s*([\d\.,]+)', val_str)
        if match:
            return match.group(1).strip()
            
    return val_str.strip()

def parse_binance_table(page):
    """
    精準定位幣安槓桿與保證金表格，在 A 欄插入 "交易所" (固定值 Binance)，
    並自動清洗數值範圍與對齊欄位。
    """
    rows_data = []
    
    table_element = page.locator("div.leverageMargin-pane table").first
    if table_element.count() == 0:
        table_element = page.locator("table").first

    if table_element.count() > 0:
        th_elements = table_element.locator("thead th, tr th").all()
        headers = ["交易所"] + [th.inner_text().strip() for th in th_elements] if len(th_elements) > 0 else []
        
        tr_elements = table_element.locator("tbody tr, tr").all()
        for tr in tr_elements:
            tds = tr.locator("td").all()
            if len(tds) >= 2:
                row_vals = ["Binance"] + [clean_range_value(td.inner_text()) for td in tds]
                
                if headers and len(headers) == len(row_vals):
                    row_dict = {headers[i]: row_vals[i] for i in range(len(headers))}
                else:
                    row_dict = {f"欄位_{i+1}": row_vals[i] for i in range(len(row_vals))}
                rows_data.append(row_dict)
                
        if rows_data:
            df = pd.DataFrame(rows_data)
            df["單位"] = "USDT"

            col_map = {}
            for c in df.columns:
                c_str = str(c).strip()
                if c_str in ["等級", "檔位"]:
                    col_map[c] = "檔位"
                elif any(k in c_str for k in ["倉位", "價值", "名義"]) or c_str == "欄位_2":
                    col_map[c] = "倉位分級"
                elif "槓桿" in c_str or c_str == "欄位_3":
                    col_map[c] = "最大槓桿"
                elif "保證金" in c_str or c_str == "欄位_4":
                    col_map[c] = "維持保證金率"
                elif "金額" in c_str or c_str == "欄位_5":
                    col_map[c] = "維持金額(USDT)"

            renamed = df.rename(columns=col_map)
            standard_cols = ["交易所", "檔位", "單位", "倉位分級", "最大槓桿", "維持保證金率", "維持金額(USDT)"]
            cols = [c for c in standard_cols if c in renamed.columns]
            return renamed[cols]

    main_content = page.locator("div.leverageMargin-pane").first
    text_lines = [clean_range_value(line) for line in main_content.inner_text().split('\n') if line.strip()]
    
    fallback_rows = [{"交易所": "Binance", "內容": line} for line in text_lines]
    return pd.DataFrame(fallback_rows)

def crawl_binance_leverage_margin(coins=None, save_excel=True):
    if coins is None:
        raw_input = input("請輸入想要查詢幣安槓桿保證金的幣種 (用空白分隔，例如: BTCUSDT ETHUSDT SOLUSDT): ").strip().upper()
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

    print(f"\n📋 [Binance 幣安] 即將依次查詢以下 {len(coins)} 個幣種: {coins}")

    url = "https://www.binance.com/zh-TC/futures/trading-parameters/perpetual/leverage-margin"

    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    excel_filename = os.path.join(data_dir, "binance_leverage_margin.xlsx")

    print("\n🚀 正在啟用 Playwright 進行幣安數據抓取（針對『幣種 永續』標籤進行完美比對）...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={'width': 1280, 'height': 800}
        )
        page = context.new_page()

        writer = pd.ExcelWriter(excel_filename, engine='openpyxl') if save_excel else None

        try:
            print(f"🔗 正在連線至: {url}")
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3000)

            results = {}
            for index, coin in enumerate(coins):
                print(f"\n--------------------------------------------------")
                print(f"🔄 正在處理第 {index+1} 個幣種: [{coin}]...")

                # 第一個幣種若是 BTCUSDT，直接抓取預設頁面
                if index == 0 and coin == "BTCUSDT":
                    print("⚡ [BTCUSDT] 為預設載入幣種，直接擷取表格！")
                else:
                    print("👉 點擊開啟幣安幣種下拉選單...")
                    
                    pane = page.locator("div.leverageMargin-pane")
                    dropdown_trigger = pane.locator("div.select-wrap, div.bn-select-trigger").first
                    dropdown_trigger.click()
                    page.wait_for_timeout(800)

                    base_coin = coin.replace('USDT', '')
                    # 清空輸入框並輸入目標幣種名稱 (例如 SOLUSDT, PEPE)
                    search_input = page.locator("input[type='text'], input[placeholder*='搜尋'], input[placeholder*='Search']").last
                    if search_input.count() > 0 and search_input.is_visible():
                        search_input.click()
                        page.keyboard.press("Control+A")
                        page.keyboard.press("Backspace")
                        print(f"⌨️ 在搜尋框中輸入 [{coin}]...")
                        page.keyboard.type(coin, delay=80)
                    else:
                        page.keyboard.type(coin, delay=80)

                    page.wait_for_timeout(1000)

                    # 廣泛與精準候選匹配文字 (含 1000PEPEUSDT 永續 等 Meme 幣命名規律)
                    target_candidates = [
                        f"{coin} 永續",
                        coin,
                        f"1000{coin} 永續",
                        f"1000{coin}",
                        f"1000{base_coin}USDT 永續",
                        f"1000{base_coin}USDT"
                    ]
                    
                    options = page.locator("div[class*='bn-select-option'], div[class*='select-option'], div[role='option']").all()
                    
                    target_clicked = False
                    # 1. 嚴格對應比對
                    for opt in options:
                        txt = opt.inner_text().strip()
                        first_line = txt.split('\n')[0].strip()
                        if txt in target_candidates or first_line in target_candidates:
                            opt.click()
                            print(f"🎯 完美對應！成功點擊選單選項 [{txt}]！")
                            target_clicked = True
                            break

                    # 2. 備援名稱包含比對 (例如 搜尋 PEPE 時匹配 1000PEPEUSDT)
                    if not target_clicked:
                        for opt in options:
                            txt = opt.inner_text().strip()
                            first_line = txt.split('\n')[0].strip()
                            if base_coin in first_line or f"1000{base_coin}" in first_line:
                                opt.click()
                                print(f"🎯 [Binance 備援] 成功點擊選單選項 [{first_line}]！")
                                target_clicked = True
                                break

                    if not target_clicked:
                        print(f"⚠️ [Binance] 未在選單中找到匹配 [{coin}] 的選項！")

                    # 等待 React 伺服器回應與 DOM 更新
                    print("⏳ 等待數據更新中...")
                    page.wait_for_timeout(3500)

                # 擷取、轉換並寫入 Excel
                print(f"✅ 正在擷取並解析 [{coin}] 表格數據...")
                df = parse_binance_table(page)
                results[coin] = df
                
                if writer:
                    sheet_name = coin[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"💾 已成功寫入獨立 Excel 工作表: [{sheet_name}]")

            if writer:
                writer.close()
                print(f"\n🎉 幣安所有幣種 ({len(coins)} 個) 數據抓取完成！檔案已儲存至: {excel_filename}")
            else:
                print(f"\n🎉 幣安所有幣種 ({len(coins)} 個) 數據抓取完成！")

            return results

        except Exception as e:
            if writer:
                writer.close()
            print(f"❌ 執行過程發生錯誤: {e}")
            return {}

        finally:
            browser.close()
            print("\n🔒 瀏覽器已安全關閉。")

if __name__ == "__main__":
    crawl_binance_leverage_margin()