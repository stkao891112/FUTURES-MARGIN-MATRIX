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
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(BASE_DIR, filename)

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

@app.route('/api/refresh', methods=['POST'])
def refresh_data():
    try:
        print("🚀 收到前端請求，準備重跑五大交易所合約保證金爬蟲 (main.py)...")
        python_exe = sys.executable
        
        # 執行 main.py
        result = subprocess.run(
            [python_exe, os.path.join(BASE_DIR, "main.py")],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )

        if result.returncode != 0:
            print(f"❌ 執行 main.py 失敗: {result.stderr}")
            return jsonify({
                "success": False,
                "error": result.stderr or "執行 main.py 時發生錯誤"
            }), 500

        # 執行 export_json.py 導出最新 JSON 並同步 index.html
        export_result = subprocess.run(
            [python_exe, os.path.join(BASE_DIR, "export_json.py")],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )

        # 讀取最新時間戳
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        json_path = os.path.join(BASE_DIR, "data", "last_updated.json")
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                ts_data = json.load(f)
                timestamp_str = ts_data.get("last_updated", timestamp_str)

        print(f"✅ 爬蟲與轉出完成！最新時間戳: {timestamp_str}")
        return jsonify({
            "success": True,
            "last_updated": timestamp_str,
            "message": "五大交易所資料更新成功！"
        })

    except Exception as e:
        print(f"❌ 處理更新請求時出錯: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/coins', methods=['GET'])
def get_coins():
    coins_path = os.path.join(BASE_DIR, "data", "coins.json")
    if os.path.exists(coins_path):
        try:
            with open(coins_path, "r", encoding="utf-8") as f:
                return jsonify({"coins": json.load(f)})
        except Exception:
            pass
    return jsonify({"coins": ["BTCUSDT", "ETHUSDT", "SOLUSDT"]})

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
        if not coin:
            return jsonify({"success": False, "error": "幣種名稱不可空白"}), 400

        if not coin.endswith('USDT'):
            coin += 'USDT'

        coins_path = os.path.join(BASE_DIR, "data", "coins.json")
        coins = ["BTCUSDT", "ETHUSDT", "SOLUSDT"]
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

        # 啟動非同步背景線程進行爬取
        t = threading.Thread(target=run_background_crawler, args=(coin,))
        t.daemon = True
        t.start()

        return jsonify({"success": True, "coins": coins, "added": coin, "message": "幣種已新增，背景正進行連線爬取！"})
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
        coins = ["BTCUSDT", "ETHUSDT", "SOLUSDT"]
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

        # 刪除 Excel 文件中該幣種 Sheet
        excel_path = os.path.join(BASE_DIR, "data", "all_exchanges_futures_margin.xlsx")
        if os.path.exists(excel_path):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(excel_path)
                if coin in wb.sheetnames:
                    del wb[coin]
                    wb.save(excel_path)
            except Exception as e:
                print("刪除 Excel 工作表提醒:", e)

        python_exe = sys.executable
        subprocess.run(
            [python_exe, os.path.join(BASE_DIR, "export_json.py")],
            cwd=BASE_DIR,
            capture_output=True
        )

        return jsonify({"success": True, "coins": coins, "deleted": coin})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    port = 5000
    print(f"🌐 FUTURES MARGIN MATRIX 伺服器啟動於: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
