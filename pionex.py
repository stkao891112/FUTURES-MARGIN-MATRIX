import asyncio
import json
import re
import sys
import os
import pandas as pd
from playwright.async_api import async_playwright

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

def parse_pionex_text_to_tiers(raw_text: str):
    """
    解析 Pionex DOM 頁面內文，轉換為標準階梯檔位字典
    """
    tiers = []
    lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
    full_text = " ".join(lines)

    # 比對格式: Lv.1 0-2,000,000 100x 0.50% 0
    pattern = re.compile(r'Lv\.?(\d+)\s+([\d,]+)\s*[-~–—]\s*([\d,]+|\u65e0\u9650|\u221e|MAX)?\s+(\d+)x\s+([\d\.]+)%\s+([\d,]+)', re.IGNORECASE)
    matches = pattern.findall(full_text)

    if matches:
        for m in matches:
            t_idx = int(m[0])
            min_v = float(m[1].replace(",", ""))
            max_v_str = m[2]
            if not max_v_str or max_v_str in ["無限", "∞", "MAX"]:
                max_v = float("inf")
            else:
                max_v = float(max_v_str.replace(",", ""))
            lev = float(m[3])
            mmr_rate = float(m[4]) / 100.0
            ded_v = float(m[5].replace(",", ""))
            
            tiers.append({
                "tier": t_idx,
                "min_amount": min_v,
                "max_amount": max_v,
                "mmr": mmr_rate,
                "deduction": ded_v,
                "max_leverage": lev
            })
    else:
        # 彈性分區塊解析
        pattern_flex = re.compile(r'(Lv\.?\d+)\s+(.*?)(?=(Lv\.?\d+|$))', re.DOTALL)
        blocks = pattern_flex.findall(full_text)
        for b_name, b_content, _ in blocks:
            range_match = re.search(r'([\d,]+)\s*[-~–—]\s*([\d,]+|無限|\u221e)?', b_content)
            lev_match = re.search(r'(\d+)x', b_content)
            mmr_match = re.search(r'([\d\.]+)%', b_content)
            ded_match = re.search(r'(?:0|[\d,]{2,})', b_content)

            if range_match and lev_match and mmr_match:
                min_val = float(range_match.group(1).replace(",", ""))
                max_val_str = range_match.group(2)
                if not max_val_str or max_val_str in ["無限", "∞"]:
                    max_val = float("inf")
                else:
                    max_val = float(max_val_str.replace(",", ""))

                lev = float(lev_match.group(1))
                mmr_rate = float(mmr_match.group(1)) / 100.0
                ded_val = float(ded_match.group(0).replace(",", "")) if ded_match else 0.0

                tier_idx = int(re.sub(r'\D', '', b_name))
                tiers.append({
                    "tier": tier_idx,
                    "min_amount": min_val,
                    "max_amount": max_val,
                    "mmr": mmr_rate,
                    "deduction": ded_val,
                    "max_leverage": lev
                })

    return tiers

async def _async_crawl_pionex(coins: list[str]) -> dict[str, pd.DataFrame]:
    results = {}

    async with async_playwright() as p:
        browser = None
        for ch in ["chrome", "msedge"]:
            try:
                browser = await p.chromium.launch(
                    channel=ch,
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
                )
                break
            except Exception:
                pass
        
        if not browser:
            browser = await p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
            )

        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1600, "height": 950},
            locale="zh-TW"
        )

        page = await context.new_page()
        await page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")

def get_pionex_symbol_candidates(coin_raw: str) -> tuple[list[str], str]:
    """
    動態產生 Pionex 的候選合約標的名稱。
    針對美股 / 指數標的，Pionex 網址結構會多一個 X，如 SOXLX.PERP_USDT
    """
    coin = coin_raw.upper().strip()
    if not coin.endswith("USDT"):
        standard_coin = f"{coin}USDT"
        base = coin
    else:
        standard_coin = coin
        base = coin[:-4]

    if coin.startswith("NCSK"):
        match = re.search(r'NCSK(.*?)2USD', coin)
        base = match.group(1) if match else base

    KNOWN_STOCKS = {
        "SOXL", "MSTR", "NVDA", "TSLA", "AAPL", "AMD", "MSFT", 
        "AMZN", "GOOGL", "META", "COIN", "PLTR", "ARM", "SMCI",
        "NFLX", "DIS", "BA", "INTC", "QCOM", "SPY", "QQQ"
    }

    is_stock = coin.startswith("NCSK") or base in KNOWN_STOCKS or (len(base) > 2 and base.endswith("X") and base[:-1] in KNOWN_STOCKS)

    if is_stock:
        if base.endswith("X"):
            candidates = [f"{base}.PERP_USDT", f"{base[:-1]}.PERP_USDT"]
        else:
            candidates = [f"{base}X.PERP_USDT", f"{base}.PERP_USDT"]
    else:
        # 加密貨幣：優先 {base}.PERP_USDT，若無數據備用嘗試 {base}X.PERP_USDT
        candidates = [f"{base}.PERP_USDT", f"{base}X.PERP_USDT"]

    return candidates, standard_coin

async def _async_crawl_pionex(coins: list[str]) -> dict[str, pd.DataFrame]:
    results = {}

    async with async_playwright() as p:
        browser = None
        for ch in ["chrome", "msedge"]:
            try:
                browser = await p.chromium.launch(
                    channel=ch,
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
                )
                break
            except Exception:
                pass
        
        if not browser:
            browser = await p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
            )

        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1600, "height": 950},
            locale="zh-TW"
        )

        page = await context.new_page()
        await page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")

        for coin_raw in coins:
            candidates, standard_coin = get_pionex_symbol_candidates(coin_raw)
            print(f"\n--------------------------------------------------")
            print(f"🔄 [Pionex] 正在處理 [{standard_coin}] 候選網址: {candidates}")

            df = pd.DataFrame()

            for cand in candidates:
                target_url = f"https://www.pionex.com/zh-TW/futures/{cand}/Bot"
                print(f"🔄 [Pionex] 連線網址 [{cand}] -> {target_url}")

                try:
                    await page.goto(target_url, wait_until="domcontentloaded", timeout=25000)
                    await page.wait_for_timeout(3500)

                    # 尋找並點擊 [幣種概況]
                    coin_tabs = page.locator("text='幣種概況'")
                    if await coin_tabs.count() > 0:
                        await coin_tabs.first.click()
                        await page.wait_for_timeout(1200)
                    else:
                        alt_tabs = page.locator("div, span, p").filter(has_text=re.compile(r"^幣種概況$|^概況$"))
                        if await alt_tabs.count() > 0:
                            await alt_tabs.first.click()
                            await page.wait_for_timeout(1200)

                    # 尋找並點擊 [槓桿與保證金]
                    margin_tabs = page.locator("text='槓桿與保證金'")
                    if await margin_tabs.count() > 0:
                        await margin_tabs.first.click()
                        await page.wait_for_timeout(2500)

                    body_text = await page.inner_text("body")
                    tiers = parse_pionex_text_to_tiers(body_text)

                    if tiers:
                        df_rows = []
                        for t in tiers:
                            min_fmt = f"{int(t['min_amount']):,}"
                            max_fmt = "無上限" if t['max_amount'] == float('inf') else f"{int(t['max_amount']):,}"
                            range_str = f"{min_fmt} ~ {max_fmt}"
                            mmr_str = f"{t['mmr']*100:.2f}%"
                            ded_str = f"{int(t['deduction']):,}" if t['deduction'] > 0 else "-"
                            lev_str = f"{int(t['max_leverage'])}x"

                            df_rows.append({
                                "交易所": "Pionex (派網)",
                                "檔位": f"檔位 {t['tier']}",
                                "單位": "USDT",
                                "名義價值範圍": range_str,
                                "維持保證金率": mmr_str,
                                "維持金額(USDT)": ded_str,
                                "最高槓桿": lev_str
                            })
                        df = pd.DataFrame(df_rows)
                        print(f"✅ [Pionex] 成功擷取 [{standard_coin}] 階梯檔位 ({len(df)} 檔) (使用網址 {cand})")
                        break
                    else:
                        print(f"⚠️ [Pionex] 網址 [{cand}] 無檔位數據，嘗試下一個網址...")

                except Exception as e:
                    print(f"❌ [Pionex] 嘗試 [{cand}] 失敗: {e}")

            results[standard_coin] = df

        await browser.close()

    return results

def crawl_pionex_position_tiers(coins: list[str], save_excel=False) -> dict[str, pd.DataFrame]:
    """
    Pionex 派網合約階梯檔位爬蟲進入點
    """
    return asyncio.run(_async_crawl_pionex(coins))

if __name__ == "__main__":
    test_coins = ["BTCUSDT", "ETHUSDT", "SOLUSDT"]
    res = crawl_pionex_position_tiers(test_coins)
    print(json.dumps(res, indent=2, ensure_ascii=False))
