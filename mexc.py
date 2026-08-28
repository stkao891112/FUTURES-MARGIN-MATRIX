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

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def get_mexc_symbol_and_unit(symbol_input: str):
    """
    將輸入的幣種字串（如 BTC, BTCUSDT, BTC_USDT）轉換為 MEXC 網址格式 (BTC_USDT) 與基礎單位 (BTC)
    """
    sym = symbol_input.upper().strip()
    if sym.endswith("-USDT"):
        sym = sym.replace("-USDT", "_USDT")
    elif sym.endswith("USDT") and "_" not in sym:
        base = sym[:-4]
        sym = f"{base}_USDT"
    elif "_" not in sym:
        sym = f"{sym}_USDT"

    base_unit = sym.split("_")[0]
    return sym, base_unit

def parse_mexc_quantity(val_str: str) -> str:
    """
    清理 MEXC 檔位數量欄位，僅萃取括號內上限數值
    例如 "0 張~50,000 張(0 BTC~5 BTC)" -> "5"
    "50,001 張~310,000 張(5.0001 BTC~31 BTC)" -> "31"
    "17,500,001 張~19,000,000 張(1,750.0001 BTC~1,900 BTC)" -> "1,900"
    """
    if not val_str:
        return ""
    val_str = str(val_str).strip()

    # 1. 先嘗試抓取括號內的內容
    match = re.search(r'\((.*?)\)', val_str)
    target_str = match.group(1) if match else val_str

    # 2. 尋找分隔符 (~, ～, -) 拆分，取右側上限
    for sep in ['~', '～', '-']:
        if sep in target_str:
            target_str = target_str.split(sep)[-1].strip()
            break

    # 3. 提取數字（含逗號與小數點）
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

def parse_mexc_table(page, symbol_input: str):
    """
    解析 MEXC 倉位風險限額檔位表格，對齊 7 大標準欄位
    """
    rows_data = []
    mexc_symbol, base_unit = get_mexc_symbol_and_unit(symbol_input)

    # 尋找頁面中的表格
    tables = page.locator("table").all()
    target_table = None

    # MEXC Risk limit 表格通常包含 "檔位" 與 "最高槓桿"
    for table in tables:
        text = table.inner_text()
        if "檔位" in text or "Tier" in text or "最高槓桿" in text:
            target_table = table
            break

    if not target_table and tables:
        target_table = tables[-1]

    if target_table:
        tr_elements = target_table.locator("tr").all()
        for tr in tr_elements:
            tds = tr.locator("td, th").all()
            if len(tds) >= 4:
                cell_texts = [td.inner_text().strip().replace('\n', ' ') for td in tds]
                
                # 排除 Header 列
                if "檔位" in cell_texts[0] or "Tier" in cell_texts[0] or "方向" in cell_texts[0]:
                    continue
                
                tier = cell_texts[0]
                qty = parse_mexc_quantity(cell_texts[1]) if len(cell_texts) > 1 else ""
                max_lev = cell_texts[2] if len(cell_texts) > 2 else ""
                maint_margin = cell_texts[3] if len(cell_texts) > 3 else ""

                # 確保槓桿包含 x 後綴
                if max_lev and not max_lev.endswith('x') and not max_lev.endswith('X'):
                    if max_lev.isdigit():
                        max_lev = f"{max_lev}x"

                rows_data.append({
                    "交易所": "MEXC",
                    "檔位": tier,
                    "單位": base_unit,
                    "倉位分級": qty,
                    "最大槓桿": max_lev,
                    "維持保證金率": maint_margin,
                    "維持金額(USDT)": "-"
                })

    if rows_data:
        return pd.DataFrame(rows_data)

    return pd.DataFrame()

def crawl_mexc_position_tiers(coins=None, save_excel=True):
    if coins is None:
        import sys
        if len(sys.argv) > 1:
            coins = [c.upper() for c in sys.argv[1:] if c.strip()]
        else:
            raw_input = input("請輸入想要查詢 MEXC 檔位的幣種 (用空白分隔，例如: BTC ETH SOL): ").strip().upper()
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

    print(f"\n📋 [MEXC] 即將依次查詢以下 {len(coins)} 個幣種: {coins}")

    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    excel_filename = os.path.join(data_dir, "mexc_position_tier.xlsx")
    results = {}

    print("\n🚀 正在啟用 Playwright 進行 MEXC 批次抓取...")

    with sync_playwright() as p:
        browser = None
        for ch in ["chrome", "msedge"]:
            try:
                browser = p.chromium.launch(
                    channel=ch,
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
                )
                print(f"🌐 [MEXC] 成功調用系統瀏覽器通道: {ch}")
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
            viewport={'width': 1280, 'height': 800}
        )

        writer = pd.ExcelWriter(excel_filename, engine='openpyxl') if save_excel else None

        try:
            for coin in coins:
                mexc_symbol, base_unit = get_mexc_symbol_and_unit(coin)
                url = f"https://www.mexc.com/zh-TW/futures/information/risk_limit/{mexc_symbol}"
                print(f"\n--------------------------------------------------")
                print(f"🔄 正在連線至 MEXC 幣種 [{mexc_symbol}]: {url}")

                page = context.new_page()
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=30000)
                    page.wait_for_timeout(3500)

                    # 自動點擊切換選單至「標的幣種」
                    try:
                        dropdown_triggers = page.locator("div[class*='select']").all()
                        if dropdown_triggers:
                            dropdown_triggers[-1].click(force=True)
                            page.wait_for_timeout(800)
                            target_coin_opt = page.locator("text='標的幣種'").first
                            if target_coin_opt.count() > 0:
                                target_coin_opt.click(force=True)
                                print(f"🎯 成功將 [{coin}] 單位切換為 [標的幣種]！")
                                page.wait_for_timeout(1500)
                    except Exception as select_err:
                        print(f"⚠️ 切換「標的幣種」選單提示: {select_err}")

                    df = parse_mexc_table(page, coin)
                    standard_coin_key = f"{base_unit}USDT"
                    results[standard_coin_key] = df

                    if not df.empty:
                        print(f"✅ 成功擷取 [{standard_coin_key}] 共 {len(df)} 檔 MEXC 數據！")
                        if writer:
                            sheet_name = standard_coin_key[:31]
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"💾 已成功寫入獨立 Excel 工作表: [{sheet_name}]")
                    else:
                        print(f"⚠️ [{standard_coin_key}] 擷取失敗或無數據。")
                except Exception as ex:
                    print(f"❌ 擷取 [{coin}] 發生異常: {ex}")
                finally:
                    page.close()

            if writer:
                writer.close()
                print(f"\n🎉 MEXC 數據抓取完成！檔案已儲存至: {excel_filename}")
            else:
                print(f"\n🎉 MEXC 數據抓取完成！")

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
    crawl_mexc_position_tiers()
