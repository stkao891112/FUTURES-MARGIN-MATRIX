import os
import sys
import json
import subprocess
import threading
from datetime import datetime
from flask import Flask, send_from_directory, jsonify, request
from flask_cors import CORS

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
CORS(app)

@app.route('/')
def serve_index():
    res = send_from_directory(BASE_DIR, 'index.html')
    res.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return res

@app.route('/<path:filename>')
def serve_static(filename):
    res = send_from_directory(BASE_DIR, filename)
    if filename.endswith('.html') or filename == 'sw.js' or filename.endswith('.json'):
        res.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return res

@app.route('/api/last_updated', methods=['GET'])
def get_last_updated():
    json_path = os.path.join(BASE_DIR, "data", "last_updated.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return jsonify(data)
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    
    # Fallback to file modification time
    excel_path = os.path.join(BASE_DIR, "data", "all_exchanges_futures_margin.xlsx")
    if os.path.exists(excel_path):
        mtime = os.path.getmtime(excel_path)
        mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        return jsonify({"last_updated": mtime_str})
    
    return jsonify({"last_updated": "尚未更新"})

def generate_temp_tiers_and_compare():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(BASE_DIR, "data", "all_exchanges_futures_margin.xlsx")
    tiers_path = os.path.join(BASE_DIR, "data", "tiers.json")

    old_data = {}
    if os.path.exists(tiers_path):
        try:
            with open(tiers_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
        except Exception:
            old_data = {}

    if not os.path.exists(excel_path):
        return False, [], {}

    from export_json import parse_tier_num, parse_number, parse_mmr
    excel_file = pd.ExcelFile(excel_path)
    new_data = {}
    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_data["_last_updated"] = timestamp_str

    for sheet in excel_file.sheet_names:
        coin = sheet.upper()
        df = pd.read_excel(excel_file, sheet_name=sheet)
        new_data[coin] = {
            "Binance": [], "Bitget": [], "Bybit": [], "OKX": [], "MEXC": [], "BingX": [], "Pionex": [], "Hyperliquid": []
        }
        if coin == "BTCUSDT":
            new_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.0125, "deduction": 0, "maxLev": 50}]
        elif coin == "ETHUSDT":
            new_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.02, "deduction": 0, "maxLev": 50}]
        elif coin == "SOLUSDT":
            new_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.025, "deduction": 0, "maxLev": 25}]
        else:
            new_data[coin]["Hyperliquid"] = [{"tier": "固定", "limit": 999999999, "mmr": 0.05, "deduction": 0, "maxLev": 20}]

        for _, row in df.iterrows():
            ex = str(row.get("交易所", "")).strip()
            if "Binance" in ex: ex_key = "Binance"
            elif "Bitget" in ex: ex_key = "Bitget"
            elif "Bybit" in ex: ex_key = "Bybit"
            elif "OKX" in ex: ex_key = "OKX"
            elif "MEXC" in ex: ex_key = "MEXC"
            elif "BingX" in ex: ex_key = "BingX"
            elif "Pionex" in ex: ex_key = "Pionex"
            else: continue

            item = {
                "tier": parse_tier_num(row.get("檔位", 1)),
                "mmr": parse_mmr(row.get("維持保證金率", 0)),
                "deduction": parse_number(row.get("維持金額(USDT)", 0)),
                "maxLev": parse_number(row.get("最大槓桿", 100))
            }
            limit_val = parse_number(row.get("倉位分級", 0))
            if ex_key in ["OKX", "MEXC"]: item["limitQty"] = limit_val
            else: item["limit"] = limit_val
            new_data[coin][ex_key].append(item)

    temp_path = os.path.join(BASE_DIR, "data", "tiers_temp.json")
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(new_data, f, ensure_ascii=False, indent=2)

    changes = []
    all_coins = set(list(old_data.keys()) + list(new_data.keys()))
    all_coins.discard("_last_updated")

    for coin in sorted(all_coins):
        old_coin = old_data.get(coin, {})
        new_coin = new_data.get(coin, {})
        all_exchanges = set(list(old_coin.keys()) + list(new_coin.keys()))
        for ex in sorted(all_exchanges):
            old_list = old_coin.get(ex, [])
            new_list = new_coin.get(ex, [])

            if old_list != new_list:
                item_diff = {
                    "coin": coin,
                    "exchange": ex,
                    "oldCount": len(old_list),
                    "newCount": len(new_list),
                    "details": []
                }
                if not old_list and new_list:
                    item_diff["summary"] = f"新增 [{coin}] {ex} 數據 (共 {len(new_list)} 檔)"
                elif old_list and not new_list:
                    item_diff["summary"] = f"未擷取到數據 (原 {len(old_list)} 檔 ➔ 0 檔)"
                elif len(old_list) != len(new_list):
                    item_diff["summary"] = f"檔位總數變更 ({len(old_list)} 檔 ➔ {len(new_list)} 檔)"
                else:
                    diff_details = []
                    for i in range(min(len(old_list), len(new_list))):
                        o = old_list[i]
                        n = new_list[i]
                        if o != n:
                            diffs = []
                            if o.get("mmr") != n.get("mmr"):
                                diffs.append(f"MMR: {(o.get('mmr',0)*100):.2f}% ➔ {(n.get('mmr',0)*100):.2f}%")
                            if o.get("maxLev") != n.get("maxLev"):
                                diffs.append(f"槓桿: {o.get('maxLev')}x ➔ {n.get('maxLev')}x")
                            if o.get("deduction") != n.get("deduction"):
                                diffs.append(f"扣減額: {o.get('deduction')} ➔ {n.get('deduction')}")
                            diff_details.append(f"第 {i+1} 檔 (" + ", ".join(diffs) + ")")
                    item_diff["details"] = diff_details
                    item_diff["summary"] = f"檔位細節變更 (" + ", ".join(diff_details[:2]) + ")"
                changes.append(item_diff)

    has_changes = len(changes) > 0
    return has_changes, changes, new_data

@app.route('/api/refresh', methods=['POST'])
def refresh_data():
    try:
        req = request.get_json() or {}
        action = req.get('action', 'audit')

        if action == 'apply':
            # 套用審核通過的暫存檔數據 (支援選擇性部分套用)
            temp_path = os.path.join(BASE_DIR, "data", "tiers_temp.json")
            tiers_path = os.path.join(BASE_DIR, "data", "tiers.json")
            selected_items = req.get('selected', None) # 格式: [{"coin": "BTCUSDT", "exchange": "BingX"}]

            if os.path.exists(temp_path):
                with open(temp_path, "r", encoding="utf-8") as f:
                    temp_data = json.load(f)

                # 讀取現有 tiers.json 作為基準
                current_data = {}
                if os.path.exists(tiers_path):
                    try:
                        with open(tiers_path, "r", encoding="utf-8") as f:
                            current_data = json.load(f)
                    except Exception:
                        current_data = {}

                # 若用戶勾選了部分項目，僅合併用戶勾選的幣種+交易所數據
                if selected_items is not None and isinstance(selected_items, list):
                    updated_count = 0
                    for sel in selected_items:
                        coin = str(sel.get('coin', '')).strip().upper()
                        ex = str(sel.get('exchange', '')).strip()
                        if coin in temp_data and ex in temp_data[coin]:
                            if coin not in current_data:
                                current_data[coin] = {
                                    "Binance": [], "Bitget": [], "Bybit": [], "OKX": [], "MEXC": [], "BingX": [], "Pionex": [], "Hyperliquid": []
                                }
                            current_data[coin][ex] = temp_data[coin][ex]
                            updated_count += 1
                    target_data = current_data
                    print(f"🎯 精確套用: 共合併用戶勾選的 {updated_count} 項交易所數據！")
                else:
                    target_data = temp_data
                    print("🎯 全數套用: 合併所有變動數據！")

                timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                target_data["_last_updated"] = timestamp_str

                with open(tiers_path, "w", encoding="utf-8") as f:
                    json.dump(target_data, f, ensure_ascii=False, indent=2)

                last_updated_path = os.path.join(BASE_DIR, "data", "last_updated.json")
                with open(last_updated_path, "w", encoding="utf-8") as f:
                    json.dump({"last_updated": timestamp_str}, f, ensure_ascii=False, indent=2)

                # 同步寫入 index.html
                html_path = os.path.join(BASE_DIR, "index.html")
                if os.path.exists(html_path):
                    with open(html_path, "r", encoding="utf-8") as f:
                        html_content = f.read()
                    tiers_json_str = json.dumps(target_data, ensure_ascii=False, indent=2)
                    updated_html = re.sub(
                        r'let EXCHANGE_TIERS = \{.*?\};',
                        f"let EXCHANGE_TIERS = {tiers_json_str};",
                        html_content,
                        flags=re.DOTALL
                    )
                    if 'let LAST_UPDATED =' in updated_html:
                        updated_html = re.sub(
                            r'let LAST_UPDATED = ".*?";',
                            f'let LAST_UPDATED = "{timestamp_str}";',
                            updated_html
                        )
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(updated_html)

                print(f"✅ 已成功套用審核通過的新數據！時間戳: {timestamp_str}")
                return jsonify({"success": True, "message": "選擇的數據已成功套用！", "last_updated": timestamp_str})
            else:
                return jsonify({"success": False, "error": "暫存資料不存在，請重新進行數據巡檢"}), 400

        elif action == 'cancel':
            temp_path = os.path.join(BASE_DIR, "data", "tiers_temp.json")
            if os.path.exists(temp_path):
                try: os.remove(temp_path)
                except Exception: pass
            print("ℹ️ 用戶取消套用新數據。")
            return jsonify({"success": True, "message": "已取消更新，維持現有數據不變。"})

        else:
            # 預設動作: 巡檢比對 (audit)
            print("🚀 收到前端請求，準備執行 7 大交易所數據對比巡檢...")
            python_exe = sys.executable
            result = subprocess.run(
                [python_exe, os.path.join(BASE_DIR, "main.py")],
                cwd=BASE_DIR,
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            if result.returncode != 0:
                print(f"❌ 執行 main.py 失敗: {result.stderr}")
                return jsonify({"success": False, "error": result.stderr or "執行 main.py 時發生錯誤"}), 500

            has_changes, changes, new_data = generate_temp_tiers_and_compare()
            timestamp_str = new_data.get("_last_updated", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            print(f"🔍 [巡檢完成] 檢測到數據是否有變動: {has_changes} (共 {len(changes)} 項異動)")
            return jsonify({
                "success": True,
                "has_changes": has_changes,
                "changes": changes,
                "last_updated": timestamp_str
            })

    except Exception as e:
        print(f"❌ 處理更新請求時出錯: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

DEFAULT_COINS = ["BTCUSDT", "ETHUSDT", "SOLUSDT", "HYPEUSDT", "DOGEUSDT", "XRPUSDT", "ADAUSDT", "SOXLUSDT", "MSTRUSDT"]

@app.route('/api/coins', methods=['GET'])
def get_coins():
    coins_path = os.path.join(BASE_DIR, "data", "coins.json")
    if os.path.exists(coins_path):
        try:
            with open(coins_path, "r", encoding="utf-8") as f:
                return jsonify({"coins": json.load(f)})
        except Exception:
            pass
    return jsonify({"coins": DEFAULT_COINS})

def run_background_crawler(coin):
    try:
        python_exe = sys.executable
        print(f"🚀 [背景爬蟲啟動] 開始專項爬取新增幣種 [{coin}]...")
        res1 = subprocess.run(
            [python_exe, os.path.join(BASE_DIR, "main.py"), coin],
            cwd=BASE_DIR,
            capture_output=True
        )
        if res1.stdout:
            print("main.py output:\n", res1.stdout.decode('utf-8', errors='ignore')[-500:])
        if res1.stderr:
            print("main.py error:\n", res1.stderr.decode('utf-8', errors='ignore')[-500:])

        res2 = subprocess.run(
            [python_exe, os.path.join(BASE_DIR, "export_json.py")],
            cwd=BASE_DIR,
            capture_output=True
        )
        if res2.stdout:
            print("export_json.py output:\n", res2.stdout.decode('utf-8', errors='ignore')[-500:])

        # 標記全流程 export 寫入完畢
        status_file = os.path.join(BASE_DIR, "data", "crawl_status.json")
        try:
            with open(status_file, "r", encoding="utf-8") as f:
                sdata = json.load(f)
            if coin not in sdata: sdata[coin] = {}
            sdata[coin]["_export_finished"] = True
            with open(status_file, "w", encoding="utf-8") as f:
                json.dump(sdata, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        print(f"✅ [背景爬蟲完成] 新增幣種 [{coin}] 數據已成功導出寫入 JSON！")
    except Exception as e:
        print(f"❌ [背景爬蟲失敗] {e}")

@app.route('/api/coins', methods=['POST'])
def add_coin():
    try:
        req = request.get_json() or {}
        coin = str(req.get('coin', '')).strip().upper()
        asset_type = str(req.get('assetType', 'crypto')).strip().lower()

        if not coin:
            return jsonify({"success": False, "error": "幣種名稱不可空白"}), 400

        if not coin.endswith('USDT'):
            coin += 'USDT'

        coins_path = os.path.join(BASE_DIR, "data", "coins.json")
        coins = list(DEFAULT_COINS)
        if os.path.exists(coins_path):
            try:
                with open(coins_path, "r", encoding="utf-8") as f:
                    coins = json.load(f)
            except Exception:
                pass

        if coin not in coins:
            coins.append(coin)
            os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
            with open(coins_path, "w", encoding="utf-8") as f:
                json.dump(coins, f, ensure_ascii=False, indent=2)

        # 儲存標的種類 (assetType: crypto 或 stock) 至 data/coin_types.json
        types_path = os.path.join(BASE_DIR, "data", "coin_types.json")
        types_data = {}
        if os.path.exists(types_path):
            try:
                with open(types_path, "r", encoding="utf-8") as f:
                    types_data = json.load(f)
            except Exception:
                pass
        types_data[coin] = asset_type
        base_unit = coin[:-4] if coin.endswith('USDT') else coin
        types_data[base_unit] = asset_type
        os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
        with open(types_path, "w", encoding="utf-8") as f:
            json.dump(types_data, f, ensure_ascii=False, indent=2)

        # 啟動非同步背景線程進行爬取
        t = threading.Thread(target=run_background_crawler, args=(coin,))
        t.daemon = True
        t.start()

        return jsonify({"success": True, "coins": coins, "added": coin, "assetType": asset_type, "message": "幣種與標的種類已新增，背景正進行連線爬取！"})
    except Exception as e:
        print(f"❌ 新增幣種失敗: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/coins/status/<coin>', methods=['GET'])
def get_coin_crawl_status(coin):
    coin = coin.strip().upper()
    if not coin.endswith('USDT'):
        coin += 'USDT'
    status_file = os.path.join(BASE_DIR, "data", "crawl_status.json")
    if os.path.exists(status_file):
        try:
            with open(status_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return jsonify({"coin": coin, "exchanges": data.get(coin, {})})
        except Exception:
            pass
    return jsonify({"coin": coin, "exchanges": {}})

@app.route('/api/coins', methods=['DELETE'])
def delete_coin():
    try:
        req = request.get_json() or {}
        coin = str(req.get('coin', '')).strip().upper()
        if not coin:
            return jsonify({"success": False, "error": "請指定要刪除的幣種"}), 400

        coins_path = os.path.join(BASE_DIR, "data", "coins.json")
        coins = list(DEFAULT_COINS)
        if os.path.exists(coins_path):
            try:
                with open(coins_path, "r", encoding="utf-8") as f:
                    coins = json.load(f)
            except Exception:
                pass

        if coin in coins:
            coins.remove(coin)
            with open(coins_path, "w", encoding="utf-8") as f:
                json.dump(coins, f, ensure_ascii=False, indent=2)

        # 依據使用者需求：刪除幣種時，僅從下拉選單 (coins.json) 中移除，後台 Excel 與 JSON 中的歷史數據完整保留不刪除！

        return jsonify({"success": True, "coins": coins, "deleted": coin})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

import socket

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

if __name__ == '__main__':
    port = 5000
    local_ip = get_local_ip()
    print("=" * 60)
    print("🌐 FUTURES MARGIN MATRIX 伺服器已成功啟動！")
    print(f"   💻 電腦本機網址: http://localhost:{port}")
    print(f"   📱 手機/局域網網址: http://{local_ip}:{port} (手機需與電腦連至同一 Wi-Fi)")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=False)
