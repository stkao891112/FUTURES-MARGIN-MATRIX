import os
from playwright.sync_api import sync_playwright
import pandas as pd

def clean_range_upper_bound(val_str: str) -> str:
    """
    清理範圍字串，僅保留上限數值（例如 "0~200,000" -> "200,000", "200,000~1,000,000" -> "1,000,000"）
    """
    if not isinstance(val_str, str):
        return val_str
    
    val_str = val_str.strip()
    
    for sep in ['~', '～', '-']:
        if sep in val_str:
            parts = val_str.split(sep)
            if len(parts) > 1 and parts[-1].strip():
                return parts[-1].strip()

    return val_str

def parse_bitget_table(page):
    """
    精準定位 Bitget 倉位檔位表格，依據列 (Row) 與欄 (Cell) 提取數據，
    確保檔位 1 填入 Row 2，並完美對齊 4 大標題。
    """
    rows_data = []
    headers = ["檔位", "價值 (USDT)", "槓桿", "檔位維持保證金率"]
    
    # 優先搜尋標準 HTML <table> 標籤與 <tr> 列
    table_element = page.locator("table").first
    if table_element.count() > 0:
        tr_elements = table_element.locator("tr").all()
        for tr in tr_elements:
            tds = tr.locator("td").all()
            if len(tds) >= 4:
                row_vals = [td.inner_text().strip() for td in tds[:4]]
                if not any(h in row_vals[0] for h in headers):
                    rows_data.append({
                        "交易所": "Bitget",
                        "檔位": row_vals[0],
                        "單位": "USDT",
                        "倉位分級": clean_range_upper_bound(row_vals[1]),
                        "最大槓桿": row_vals[2],
                        "維持保證金率": row_vals[3],
                        "維持金額(USDT)": "-"
                    })
        if rows_data:
            return pd.DataFrame(rows_data)

    # 備援 DOM 切割機制
    main_content = page.locator("main, article, div[class*='container']").first
    text_lines = [line.strip() for line in main_content.inner_text().split('\n') if line.strip()]
    
    data_tokens = [line for line in text_lines if not any(h in line for h in headers) and "倉位檔位介紹" not in line]
    
    parsed_rows = []
    if len(data_tokens) >= 4:
        for i in range(0, len(data_tokens) - len(data_tokens) % 4, 4):
            parsed_rows.append({
                "交易所": "Bitget",
                "檔位": data_tokens[i],
                "單位": "USDT",
                "倉位分級": clean_range_upper_bound(data_tokens[i+1]),
                "最大槓桿": data_tokens[i+2],
                "維持保證金率": data_tokens[i+3],
                "維持金額(USDT)": "-"
            })
            
    return pd.DataFrame(parsed_rows) if parsed_rows else pd.DataFrame({"交易所": ["Bitget"]*len(text_lines), "原始資料": text_lines})

def crawl_bitget_position_tiers(coins=None, save_excel=True):
    if coins is None:
        raw_input = input("請輸入想要查詢 Bitget 檔位的幣種 (用空白分隔，例如: BTCUSDT ETHUSDT SOLUSDT): ").strip().upper()
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

    print(f"\n📋 [Bitget] 即將依次查詢以下 {len(coins)} 個幣種: {coins}")

    url = "https://www.bitget.com/zh-TC/futures/introduction/position-tier"

    # 使用 Raw String 精準定義你提供的 Bitget 幣種選單 Selector
    exact_dropdown_selector = r"#root > div > div > div.pc\:flex.pad\:flex.air\:flex.block\:phone > div.pc\:w-full.flex-1.px-\[40px\].pad\:px-\[32px\].phone\:px-\[16px\].overflow-auto.break-words > div > div.mt-\[32px\].mb-\[40px\].flex.items-center.justify-between.flex-1.phone\:mb-\[unset\] > div.flex.gap-\[12px\] > div:nth-child(3) > div.bit-select.bit-select-medium.bit-select-round.css-1oxbkxq.bit-select-single.bit-select-show-arrow.bit-select-show-search > div > span.bit-select-selection-item"

    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    excel_filename = os.path.join(data_dir, "bitget_position_tier.xlsx")

    print("\n🚀 正在啟用 Playwright 進行 Bitget 批次抓取...")

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
            is_first_coin = True

            for coin in coins:
                print(f"\n--------------------------------------------------")
                print(f"🔄 正在處理幣種: [{coin}]...")

                # 第一個幣種為 BTCUSDT 時，直接抓取預設頁面
                if is_first_coin and coin == "BTCUSDT":
                    print("⚡ [BTCUSDT] 為預設載入幣種，無需切換選單，直接擷取！")
                    is_first_coin = False
                else:
                    is_first_coin = False
                    
                    print("👉 使用精準 Selector 開啟 Bitget 合約下拉選單...")
                    dropdown = page.locator(exact_dropdown_selector)

                    if dropdown.count() > 0:
                        dropdown.scroll_into_view_if_needed()
                        page.wait_for_timeout(500)
                        dropdown.click()
                    else:
                        print("⚠️ 精準 Selector 未直接命中，嘗試點擊第三個選單容器...")
                        page.locator("div.flex.gap-\\[12px\\] > div:nth-child(3)").click()

                    page.wait_for_timeout(1000)

                    # 聚焦輸入框並輸入幣種
                    search_input = page.locator("input[type='text'], input[placeholder*='搜尋'], input[placeholder*='Search']").last
                    
                    if search_input.count() > 0 and search_input.is_visible():
                        print(f"⌨️ 在搜尋框中清空並輸入 [{coin}]...")
                        search_input.fill("")
                        search_input.type(coin)
                    else:
                        print(f"⌨️ 直接發送鍵盤輸入 [{coin}]...")
                        page.keyboard.type(coin)

                    page.wait_for_timeout(800)

                    # 選擇目標幣種選項
                    option_item = page.locator(f"text={coin}").last
                    if option_item.count() > 0 and option_item.is_visible():
                        option_item.click()
                        print(f"🎯 成功點擊選單選項 [{coin}]！")
                    else:
                        page.keyboard.press("Enter")
                        print(f"⌨️ 按下 Enter 確認選取 [{coin}]！")

                    page.wait_for_timeout(3000)

                # 擷取資料並寫入 Excel
                print(f"✅ 正在擷取並解析 [{coin}] 的檔位表格...")
                df = parse_bitget_table(page)
                results[coin] = df
                
                if writer:
                    sheet_name = coin[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"💾 已成功寫入獨立 Excel 工作表: [{sheet_name}]")

            if writer:
                writer.close()
                print(f"\n🎉 所有幣種 ({len(coins)} 個) 抓取完成！檔案已成功儲存至: {excel_filename}")
            else:
                print(f"\n🎉 所有幣種 ({len(coins)} 個) 抓取完成！")

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
    crawl_bitget_position_tiers()