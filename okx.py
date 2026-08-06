import os
import re
from playwright.sync_api import sync_playwright
import pandas as pd

def parse_okx_quantity(val_str: str) -> str:
    """
    清理 OKX 張數/數量欄位，僅萃取括號內上限數值（例如 "0.00 ~ 1,000.00 (0.0000 BTC - 10.0000 BTC)" -> "10"）
    """
    if not val_str:
        return ""
    val_str = str(val_str).strip()
    
    match = re.search(r'\((.*?)\)', val_str)
    target_str = match.group(1) if match else val_str

    for sep in ['-', '~', '～']:
        if sep in target_str:
            target_str = target_str.split(sep)[-1].strip()
            break

    num_match = re.search(r'[\d\.,]+', target_str)
    if num_match:
        raw_num = num_match.group(0).replace(',', '')
        try:
            float_val = float(raw_num)
            if float_val.is_integer():
                return f"{int(float_val):,}"
            else:
                return f"{float_val:,}"
        except ValueError:
            return num_match.group(0)

    return target_str

def get_base_coin(symbol: str) -> str:
    """
    從交易對符號提取基礎幣種單位 (例如 BTCUSDT -> BTC)
    """
    symbol = symbol.upper().strip()
    for suffix in ["-USDT-SWAP", "_USDT", "USDT"]:
        if symbol.endswith(suffix):
            return symbol[:-len(suffix)]
    return symbol

def parse_okx_table(page, coin: str):
    """
    解析 OKX 倉位檔位表格，對齊 7 大標準欄位。
    """
    rows_data = []
    base_unit = get_base_coin(coin)
    
    table_element = page.locator("table, div.okui-table").first
    if table_element.count() > 0:
        tr_elements = table_element.locator("tbody tr, tr").all()
        for tr in tr_elements:
            tds = tr.locator("td").all()
            if len(tds) >= 4:
                cell_texts = [td.inner_text().strip() for td in tds]
                # 排除 Header 列
                if "檔位" in cell_texts[0] or "Tier" in cell_texts[0]:
                    continue
                
                tier = cell_texts[0]
                qty = parse_okx_quantity(cell_texts[1]) if len(cell_texts) > 1 else ""
                maint_margin = cell_texts[2] if len(cell_texts) > 2 else ""
                # 最高可用槓桿倍數通常在最後一欄
                max_lev = cell_texts[-1] if len(cell_texts) >= 5 else (cell_texts[3] if len(cell_texts) > 3 else "")
                
                rows_data.append({
                    "交易所": "OKX",
                    "檔位": tier,
                    "單位": base_unit,
                    "倉位分級": qty,
                    "最大槓桿": max_lev,
                    "維持保證金率": maint_margin,
                    "維持金額(USDT)": "-"
                })
        if rows_data:
            return pd.DataFrame(rows_data)

    # 備援解析
    main_content = page.locator("main, article, div[class*='container']").first
    text_lines = [line.strip() for line in main_content.inner_text().split('\n') if line.strip()]
    return pd.DataFrame({"交易所": ["OKX"]*len(text_lines), "原始資料": text_lines})

def crawl_okx_position_tiers(coins=None, save_excel=True):
    if coins is None:
        raw_input = input("請輸入想要查詢 OKX 檔位的幣種 (用空白分隔，例如: BTCUSDT ETHUSDT SOLUSDT): ").strip().upper()
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

    print(f"\n📋 [OKX] 即將依次查詢以下 {len(coins)} 個幣種: {coins}")

    url = "https://www.okx.com/zh-hant/trade-market/position/swap"
    dropdown_selector = r"#root > div > div > section.market-info-container > div > section.info-query > div:nth-child(4) > div > div.okui-select-value-box > div > div"

    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    excel_filename = os.path.join(data_dir, "okx_position_tier.xlsx")

    print("\n🚀 正在啟用 Playwright 進行 OKX 批次抓取...")

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

                # 第一個幣種若是 BTCUSDT，直接擷取預設頁面
                if is_first_coin and (coin == "BTCUSDT" or coin == "BTC"):
                    print("⚡ [BTCUSDT] 為預設載入幣種，無需切換選單，直接擷取！")
                    is_first_coin = False
                else:
                    is_first_coin = False
                    try:
                        print("👉 開啟 OKX 合約幣種下拉選單...")
                        
                        # 定位 OKX 第四個下拉選單（幣種選單）
                        dropdown_trigger = page.locator(dropdown_selector)
                        if dropdown_trigger.count() == 0:
                            dropdown_trigger = page.locator("section.info-query > div:nth-child(4)").first
                        if dropdown_trigger.count() == 0:
                            dropdown_trigger = page.locator("div.info-query div.okui-select-value-box").last

                        dropdown_trigger.scroll_into_view_if_needed()
                        page.wait_for_timeout(300)
                        dropdown_trigger.click(force=True)
                        print(f"🎯 成功點擊 OKX 幣種下拉選單！")
                        page.wait_for_timeout(1000)

                        # 定位浮動選單容器 (Popover / Dropdown Container)
                        popup = page.locator("div.okui-select-dropdown, div.okui-popover, div.okui-select-popup-container, div[class*='select-popup']").last
                        
                        # 尋找浮動選單內的搜尋框
                        search_input = popup.locator("input[type='text'], input").first
                        if search_input.count() > 0 and search_input.is_visible():
                            search_input.click()
                            page.keyboard.press("Control+A")
                            page.keyboard.press("Backspace")
                            print(f"⌨️ 在 OKX 搜尋框中輸入 [{coin}]...")
                            search_input.type(coin, delay=100)
                        else:
                            print(f"⌨️ 直接發送全域鍵盤輸入 [{coin}]...")
                            page.keyboard.type(coin, delay=100)

                        page.wait_for_timeout(800)

                        # 核心優化：進行 100% 嚴格完全比對，徹底解決 MSTR 誤點選 HMSTR 等問題
                        base_unit = get_base_coin(coin)
                        exact_targets = [
                            f"{base_unit}-USDT-SWAP",
                            f"{base_unit}-USDT",
                            f"{base_unit} / USDT",
                            f"{base_unit}USDT",
                            base_unit
                        ]

                        options = popup.locator("div.okui-select-item, li, div[class*='select-item'], div[role='option'], div[class*='item']").all()
                        target_clicked = False
                        for opt in options:
                            txt = opt.inner_text().strip().upper()
                            first_line = txt.split('\n')[0].strip()
                            if txt in exact_targets or first_line in exact_targets:
                                opt.click(force=True)
                                print(f"🎯 [OKX] 100% 嚴格完全對應！成功點擊選單選項 [{first_line}]！")
                                target_clicked = True
                                break
                            
                            # 比對拆解後的首個代幣名稱 (例如 "MSTR-USDT-SWAP" -> "MSTR")
                            parts = [p.strip() for p in re.split(r'[\s/_\-]+', first_line) if p.strip()]
                            if parts and parts[0] == base_unit:
                                opt.click(force=True)
                                print(f"🎯 [OKX] 標的代碼完全相同，點擊選項 [{first_line}]！")
                                target_clicked = True
                                break

                        if not target_clicked:
                            print(f"⌨️ 未找到列表項，發送 Enter 鍵選取 [{coin}]...")
                            page.keyboard.press("Enter")

                        # 等待 OKX React DOM 與表格數據重載
                        print("⏳ 等待 OKX 表格數據更新中...")
                        page.wait_for_timeout(3500)

                    except Exception as ex_err:
                        print(f"⚠️ 切換幣種 [{coin}] 過程提示: {ex_err}，發送 Enter 繼續... ")
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(3500)

                # 擷取資料
                print(f"✅ 正在擷取並解析 [{coin}] 的 OKX 檔位表格...")
                df = parse_okx_table(page, coin)
                results[coin] = df

                if writer:
                    sheet_name = coin[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"💾 已成功寫入獨立 Excel 工作表: [{sheet_name}]")

            if writer:
                writer.close()
                print(f"\n🎉 OKX 數據抓取完成！檔案已儲存至: {excel_filename}")
            else:
                print(f"\n🎉 OKX 數據抓取完成！")

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
    crawl_okx_position_tiers()
