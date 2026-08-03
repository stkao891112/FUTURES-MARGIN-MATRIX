import os
import sys
import json
import subprocess
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

if __name__ == '__main__':
    port = 5000
    print(f"🌐 FUTURES MARGIN MATRIX 伺服器啟動於: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
