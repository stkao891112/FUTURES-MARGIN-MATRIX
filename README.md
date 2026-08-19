# 🚀 FUTURES MARGIN MATRIX

> **跨交易所永續合約保證金與爆倉清算計算機**
> 
> 自動化爬取並整合 **7 大衍生品交易所**（Binance, Bitget, Bybit, OKX, MEXC, BingX, Pionex）全幣種（含加密貨幣與美股/指數標的）之維持保證金率 (MMR)、階梯檔位、維持保證金扣減額與最大槓桿限制，並提供全響應式 PWA 清算計算機。

---

## 📌 1. 專案簡介 (Project Overview)

- **GitHub 儲存庫**：[stkao891112/FUTURES-MARGIN-MATRIX](https://github.com/stkao891112/FUTURES-MARGIN-MATRIX.git)
- **核心功能**：
  - **7 大交易所全自動爬蟲**：實時連線爬取最新維持保證金率與階梯檔位。
  - **加密貨幣 & 美股/指數標的支援**：支援 BTC, ETH, SOL, HYPE, DOGE, XRP, ADA 及美股標的 (SOXL, MSTR, NVDA 等)。
  - **爆倉價與清算距離比較**：精準計算全交易所清算爆倉價，並自動按安全性進行排名比對。
  - **全功能 PWA 介面**：支援手動拖曳卡片排序 (Drag & Drop)、動態同步進度條 (Progress Modal)、一鍵手機快取清除與局域網/外網 (Ngrok) 綁定。

---

## 🛠️ 2. 技術棧與架構 (Technology Stack)

| 分類 (Layer) | 技術 / 工具 (Stack) | 用途與說明 (Description) |
|---|---|---|
| **爬蟲核心 (Crawler Core)** | Python 3.10+, Playwright (Async), pandas, openpyxl | 自動化瀏覽器模擬、DOM 解析、動態選單切換、檔位提取與 Excel/JSON 寫入 |
| **反爬蟲繞過 (Stealth & Anti-Bot)** | Playwright Chrome System Channel (`channel="chrome"`), Stealth Scripts, Candidate URL Fallback | 繞過 Cloudflare 10006 機器人防護與動態 JavaScript 混淆 DOM |
| **並行處理 (Concurrency)** | Python `asyncio` & `ThreadPoolExecutor` | 7 大交易所並行連線獨立爬取，總耗時縮短至 ~60-70 秒 |
| **本地 API 伺服器 (Backend Server)** | Python Flask, Flask-CORS, subprocess, socket | 提供 `/api/refresh`, `/api/coins`, `/api/last_updated` 介面，支援跨網域 (CORS) 與局域網 IP 檢測 |
| **前端 UI (Frontend)** | HTML5, Vanilla CSS3 (Glassmorphism), Vanilla JS (ES6+) | 爆倉清算價計算、清算距離排序、卡片 HTML5 Drag & Drop 拖曳排序、PWA 離線支援 |
| **進度與快取 (Refresh & Cache)** | HTML5 Progress Modal, Service Worker (SW v9), CacheStorage, LocalStorage | 7 大交易所即時同步進度條、一鍵自動清除手機快取、自訂 API Host |

---

## 🛡️ 3. 核心技術難點與防護繞過方案 (Anti-Bot & Technical Bypassing)

### 3.1 難點 1：Cloudflare 10006 Error 與 Headless Bot 防護繞過 (Pionex & BingX)
- **問題現象**：
  當使用預設 Headless Chromium (`async_playwright().chromium.launch(headless=True)`) 連線 Pionex 交易介面或 BingX 網頁時，Cloudflare 防禦系統會直接攔截連線，回傳 `10006 Access Denied` 錯誤或拒絕載入 DOM。
- **繞過方案**：
  1. **指定系統真實 Chrome Channel**：
     使用 `playwright.chromium.launch(channel="chrome", headless=True)` 或 `channel="msedge"`。這會調用本機安裝的官方 Google Chrome 二進位檔，具備完整的 WebGL/Canvas/TLS 密碼套件指紋，能 100% 繞過 Cloudflare 的自動化機器人檢測。
  2. **隱藏 `navigator.webdriver` 標記**：
     設定 `viewport={'width': 1920, 'height': 1080}`、`user_agent` 並注入腳本將 `navigator.webdriver` 設為 `undefined`。

---

### 3.2 難點 2：Pionex (派網) 交易介面內嵌 DOM 導覽
- **問題現象**：
  Pionex 沒有公開的靜態階梯 MMR 網頁，所有階梯保證金資訊皆藏於 Bot 交易介面內。
- **繞過與解析方案**：
  - 爬蟲自動載入交易頁面 `https://www.pionex.com/zh-TW/futures/{SYMBOL}/Bot` 後，定位並點擊 **`[幣種概況]` (Coin Overview)** 標籤，接著點擊 **`[槓桿與保證金]` (Leverage & Margin)** 子頁籤。
  - 等待內嵌的階梯 DOM 表格渲染完成，透過 JavaScript DOM 選取器過濾並提取檔位、名義價值與 MMR 數值。

---

### 3.3 難點 3：美股與加密貨幣網址規則差異與雙向備用退回 (Candidate Fallback System)
- **問題現象**：
  - Pionex 加密貨幣網址為 `{BASE}.PERP_USDT` (如 `BTC.PERP_USDT`)；美股標的 (如 SOXL, MSTR, NVDA) 網址後方帶有 `X` 後綴 (如 `SOXLX.PERP_USDT`)。
  - BingX 美股標的網址為 `NCSK{BASE}2USD-USDT` (如 `NCSKSOXL2USD-USDT`)。
- **解決方案 (雙向自動備用退回)**：
  - 實作 `get_pionex_symbol_candidates()` 與 BingX 候選解析邏輯：
    - 識別標的種類（美股 vs 加密貨幣）。
    - 對於美股標的，優先嘗試 `SOXLX.PERP_USDT`；若提取失敗或 0 檔位則自動退回嘗試 `SOXL.PERP_USDT`。
    - 對於加密貨幣，優先嘗試 `BTC.PERP_USDT`；若提取失敗則自動嘗試 `BTCX.PERP_USDT`。
    - 100% 確保未來使用者新增任何全新幣種或美股時，系統皆能自動防錯與精準擷取。

---

### 3.4 難點 4：跨交易所數據單位標準化與 Dataframe 合併碰撞 (Index Collision)
- **問題現象**：
  各交易所欄位定義不同（如幣安為 USDT 名義價值、OKX/MEXC 為張數/幣數，Bitget 含扣減額）。合併 DataFrame `pd.concat()` 時若欄位名稱發生重複映射，Pandas 會拋出 `InvalidIndexError: Reindexing only valid with uniquely valued Index objects` 錯誤。
- **解決方案**：
  - 統一標準化 Dataframe 7 大欄位：`["交易所", "檔位", "單位", "名義價值範圍", "維持保證金率", "維持金額(USDT)", "最高槓桿"]`。
  - 在 `normalize_dataframe_columns()` 中調整正規表示式判斷順序（優先匹配「扣減額/金額」，再匹配「保證金率」），解決欄位名重複碰撞問題。

---

### 3.5 難點 5：手機端 PWA 快取與跨裝置 API 呼叫 (Mobile Cache & LAN/Ngrok IP Binding)
- **問題現象**：
  - 手機 Safari/PWA 具有強烈離線快取，導致更新數據後手機仍顯示舊版畫面。
  - 手機點擊「更新資料來源」提示 `server.py未啟動`，因為手機本機並未執行 `localhost:5000`。
- **解決方案**：
  1. **一鍵強制清除快取 (`checkUpdateAndClearCache`)**：
     自動解除 Service Worker 註冊、刪除 CacheStorage，帶入 `?v=timestamp` 切除標籤強制手機原生重整。
  2. **局域網 IP 檢測與自訂 API Host (`apiFetch`)**：
     - `server.py` 啟動時自動偵測並印出電腦局域網 IP (如 `http://192.168.1.xxx:5000`)。
     - 前端提供 `設定 Server IP` 彈窗，支援手機填入局域網 IP 或 Ngrok 外網映射網址（`https://xxxx.ngrok-free.app`），讓手機隨時隨地皆可遠端發起爬蟲！

---

## 📁 4. 專案檔案結構 (Repository Architecture)

```
FUTURES-MARGIN-MATRIX/
│
├── main.py                  # 爬蟲主入口 (整合 7 大交易所 async 任務、Dataframe 標準化與 Excel 生成)
├── server.py                # 本地 Flask RESTful API 伺服器 (支援 CORS、自動 LAN IP 偵測)
├── export_json.py           # Excel 轉出 JSON API (`data/tiers.json`) 與 index.html 數據直嵌
│
├── binance.py               # 幣安 (Binance) 階梯檔位爬蟲模組
├── bitget.py                # Bitget 階梯檔位爬蟲模組
├── bybit.py                 # Bybit 階梯檔位爬蟲模組
├── okx.py                   # OKX (歐易) 階梯檔位爬蟲模組
├── mexc.py                  # MEXC (抹茶) 階梯檔位爬蟲模組
├── bingx.py                 # BingX (包含美股 NCSK 格式) 階梯檔位爬蟲模組
├── pionex.py                # Pionex (派網, Chrome Channel 繞過防護 & 美股 X 後綴) 爬蟲模組
│
├── index.html               # 核心 PWA 前端網頁 (爆倉計算機、進度條 Modal、卡片拖曳排序、快取清除)
├── manifest.json            # PWA 應用設定檔
├── sw.js                    # Service Worker 離線快取控制 (PWA v9)
├── requirements.txt         # Python 依賴套件清單 (playwright, pandas, openpyxl, flask, flask-cors)
└── README.md                # 專案說明與總結文件
```

---

## 🚀 5. 快速開始 (Quick Start)

### 5.1 安裝依賴
```bash
pip install -r requirements.txt
playwright install chrome
```

### 5.2 啟動本地 API 伺服器
```bash
python server.py
```
控制台將印出：
```text
🌐 FUTURES MARGIN MATRIX 伺服器已成功啟動！
   💻 電腦本機網址: http://localhost:5000
   📱 手機/局域網網址: http://192.168.1.XXX:5000
```

### 5.3 執行全交易所並行爬蟲
```bash
python main.py
```

### 🤖 5.4 雲端每日自動爬蟲 (GitHub Actions + 數據變動比對)
本專案已整合 GitHub Actions 自動化工作流 (`.github/workflows/crawl.yml`)：
- **每天台灣時間 08:00 AM (UTC 00:00)** 自動在 GitHub 雲端 Linux 虛擬機啟動 7 大交易所 Playwright 並行爬蟲。
- **智能數據 Diff 檢查**：自動比對 `data/tiers.json` 與 Excel 總表。**僅在檢測到檔位/MMR 數據發生實質變動時，才會自動 Commit 並發佈至 GitHub Pages**；若數據無變動則保持靜默不做任何動作，避免產生無效 Commit 紀錄！

---

## 📝 6. 總結 (Conclusion)

本專案成功結合 **Playwright 自動化爬蟲、Flask 後端 API、Pandas 多源數據匯總與 Vanilla JS 現代化 Glassmorphism PWA 前端**，解決了跨交易所合約保證金規格不一、防爬蟲 Cloudflare 攔截、動態 DOM 解析以及行動端跨網域連線與快取的全方位技術挑戰。
