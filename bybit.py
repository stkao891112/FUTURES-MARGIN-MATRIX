import os
import sys
from playwright.sync_api import sync_playwright
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def parse_html_table_to_dataframe(table_element):
    """
    透過 Playwright 直接解析 HTML <table> 的 <tr> 與 <td>，確保 100% 分欄正確
    """
    rows_data = []
    headers = ["檔位", "風險限額", "最大槓桿", "維持保證金率", "維持保證金扣減額"]
    
    # 尋找表格內所有的 <tr> 列
    tr_elements = table_element.locator("tr").all()
    
    # 若找到 HTML 標準 <tr> 標籤
    if len(tr_elements) > 0:
        for tr in tr_elements:
            # 尋找該列所有的 <td> 儲存格
            tds = tr.locator("td").all()
            if len(tds) >= 5:
                # 提取儲存格文字
                row_vals = [td.inner_text().strip() for td in tds[:5]]
                rows_data.append({
                    "交易所": "Bybit",
                    "檔位": row_vals[0],
                    "單位": "USDT",
                    "倉位分級": row_vals[1],
                    "最大槓桿": row_vals[2],
                    "維持保證金率": row_vals[3],
                    "維持金額(USDT)": row_vals[4]
                })
        if rows_data:
            return pd.DataFrame(rows_data)

    # 備援機制：若 Bybit 採用 div 做 flex/grid 排版，透過通用 class 切割
    grid_rows = table_element.locator("div[class*='row']").all()
    if not grid_rows:
        # 尋找包含資料的父級 div
        grid_rows = table_element.locator("> div").all()

    for row in grid_rows:
        # 尋找子級的儲存格 div
        cells = row.locator("> div").all()
        if len(cells) >= 5:
            cell_texts = [c.inner_text().strip() for c in cells[:5]]
            # 排除標題列
            if "檔位" not in cell_texts[0] and "Tier" not in cell_texts[0]:
                rows_data.append({
                    "交易所": "Bybit",
                    "檔位": cell_texts[0],
                    "單位": "USDT",
                    "倉位分級": cell_texts[1],
                    "最大槓桿": cell_texts[2],
                    "維持保證金率": cell_texts[3],
                    "維持金額(USDT)": cell_texts[4]
                })

    if rows_data:
        return pd.DataFrame(rows_data)
    else:
        # 最終備援解析 (若 DOM 結構極度特殊)
        text_lines = [line.strip() for line in table_element.inner_text().split('\n') if line.strip()]
        filtered_lines = [line for line in text_lines if not any(h in line for h in headers)]
        
        parsed_rows = []
        for i in range(0, len(filtered_lines) - len(filtered_lines) % 5, 5):
            parsed_rows.append({
                "交易所": "Bybit",
                "檔位": filtered_lines[i],
                "單位": "USDT",
                "倉位分級": filtered_lines[i+1],
                "最大槓桿": filtered_lines[i+2],
                "維持保證金率": filtered_lines[i+3],
                "維持金額(USDT)": filtered_lines[i+4]
            })
        return pd.DataFrame(parsed_rows)

def crawl_bybit_multi_coins_to_excel(coins=None, save_excel=True):
    if coins is None:
        raw_input = input("請輸入想要查詢的幣種 (用空白分隔，例如: BTCUSDT ETHUSDT SOLUSDT): ").strip().upper()
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

    print(f"\n📋 [Bybit] 即將查詢以下 {len(coins)} 個幣種: {coins}")

    url = "https://www.bybit.com/zh-TW/announcement-info/margin-parameters/"

    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    excel_filename = os.path.join(data_dir, "bybit_margin_multi.xlsx")

    print("\n🚀 正在啟用 Playwright 進行批次抓取 (Chrome Channel 防護繞過模式)...")

    with sync_playwright() as p:
        browser = None
        for ch in ["chrome", "msedge"]:
            try:
                browser = p.chromium.launch(
                    channel=ch,
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
                )
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
            viewport={'width': 1600, 'height': 950},
            locale="zh-TW"
        )
        page = context.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")

        writer = pd.ExcelWriter(excel_filename, engine='openpyxl') if save_excel else None

        try:
            print(f"🔗 正在連線至: {url}")
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3500)

            results = {}
            current_active_coin = "BTCUSDT"

            for coin in coins:
                print(f"\n--------------------------------------------------")
                print(f"🔄 正在處理幣種: [{coin}]...")

                # 尋找頁面上的表格 Element
                table_locator = page.locator("article table").first
                if table_locator.count() == 0:
                    table_locator = page.locator("table").first

                # 特殊處理首個幣種，若 DOM 已呈現則免點擊，否則強制透過選單搜尋選取
                need_click = True
                if coin == "BTCUSDT" and current_active_coin == "BTCUSDT":
                    if table_locator.count() > 0 and len(table_locator.locator("tr").all()) > 1:
                        print("⚡ [BTCUSDT] 為預設幣種，網頁已完成渲染，直接擷取表格數據！")
                        need_click = False

                if need_click:
                    coin_anchor = page.locator("article div").filter(has_text=current_active_coin).last
                    if coin_anchor.count() == 0:
                        coin_anchor = page.locator("input, [class*='select']").first

                    if coin_anchor.count() > 0:
                        print(f"👉 點擊下拉選單 (當前顯示: {current_active_coin})...")
                        try:
                            coin_anchor.scroll_into_view_if_needed()
                            page.wait_for_timeout(300)
                            coin_anchor.click()
                        except Exception:
                            pass
                    else:
                        print("⚠️ 未能找到選單錨點，嘗試直接模擬鍵盤搜尋...")

                    page.wait_for_timeout(1000)

                    print(f"⌨️ 輸入幣種 [{coin}]...")
                    page.keyboard.type(coin)
                    page.wait_for_timeout(800)

                    # 在浮動選單清單中選擇第一個匹配項目 (首項即為 USDT 永續合約)
                    popover_options = page.locator("div[class*='option'], div[role='option'], li, div[class*='item']").filter(has_text=coin).all()
                    target_clicked = False
                    for opt in popover_options:
                        if opt.is_visible() and opt != coin_anchor:
                            try:
                                opt.click()
                                print(f"🎯 成功點擊 Bybit 選單首項 [{coin}] (USDT 永續)！")
                                target_clicked = True
                                break
                            except Exception:
                                pass

                    if not target_clicked:
                        page.keyboard.press("ArrowDown")
                        page.wait_for_timeout(200)
                        page.keyboard.press("Enter")
                        print(f"⌨️ 備援機制：按向下鍵與 Enter 選取 [{coin}] (首項 USDT 永續)！")

                    current_active_coin = coin
                    page.wait_for_timeout(3000)

                # 讀取 DOM 元素結構進行儲存格解析
                table_locator = page.locator("article table").first
                if table_locator.count() == 0:
                    table_locator = page.locator("table").first

                if table_locator.count() > 0:
                    print(f"✅ 成功找到 [{coin}] 表格 DOM，正在逐格提取數值...")
                    df = parse_html_table_to_dataframe(table_locator)
                    results[coin] = df
                    
                    if writer:
                        sheet_name = coin[:31]
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"💾 已成功將 [{coin}] 的標準對齊表格寫入獨立 Excel 工作表: [{sheet_name}]")
                else:
                    print(f"⚠️ 未能擷取到 [{coin}] 的表格資料。")

            if writer:
                writer.close()
                print(f"\n🎉 Bybit 任務完成！Excel 檔案已成功更新至: {excel_filename}")
            else:
                print(f"\n🎉 Bybit 任務完成！")

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
    crawl_bybit_multi_coins_to_excel()