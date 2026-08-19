import json
import subprocess
import os
import sys

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def compare_tiers():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    tiers_path = os.path.join(BASE_DIR, "data", "tiers.json")

    # 取得 Git HEAD 版本 (舊數據)
    try:
        old_bytes = subprocess.check_output(["git", "show", "HEAD:data/tiers.json"], stderr=subprocess.DEVNULL)
        old_data = json.loads(old_bytes.decode('utf-8', errors='ignore'))
    except Exception:
        old_data = {}

    # 取得當前磁碟版本 (新數據)
    if os.path.exists(tiers_path):
        try:
            with open(tiers_path, "r", encoding="utf-8") as f:
                new_data = json.load(f)
        except Exception:
            new_data = {}
    else:
        new_data = {}

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
                if not old_list and new_list:
                    changes.append(f"  ✨ [{coin}] -> [{ex}]: 新增成功，共擷取到 {len(new_list)} 檔階梯檔位數據")
                elif old_list and not new_list:
                    changes.append(f"  ⚠️ [{coin}] -> [{ex}]: 數據丟失 (原 {len(old_list)} 檔 ➔ 變為 0 檔)")
                elif len(old_list) != len(new_list):
                    changes.append(f"  🔄 [{coin}] -> [{ex}]: 檔位總數變更 ({len(old_list)} 檔 ➔ {len(new_list)} 檔)")
                else:
                    # 比較具體檔位細節
                    diff_details = []
                    for i in range(min(len(old_list), len(new_list))):
                        o_item = old_list[i]
                        n_item = new_list[i]
                        if o_item != n_item:
                            diff_details.append(f"第 {i+1} 檔 (MMR: {o_item.get('mmr')} ➔ {n_item.get('mmr')})")
                    detail_str = ", ".join(diff_details[:3])
                    if len(diff_details) > 3:
                        detail_str += "..."
                    changes.append(f"  📝 [{coin}] -> [{ex}]: 檔位細節變更 ({detail_str})")

    return changes

if __name__ == "__main__":
    print("============================================================")
    print(" 📊 交易所檔位數據變動精確報告 (Exchange Data Diff Report)")
    print("============================================================")
    
    diff_list = compare_tiers()
    if diff_list:
        print(f"✨ 檢測到共 {len(diff_list)} 項交易所檔位數據發生變動:\n")
        for item in diff_list:
            print(item)
        print("\n============================================================")
        # 設定 output 供 GitHub Actions 判斷
        with open(os.environ.get('GITHUB_OUTPUT', 'github_output.txt'), 'a', encoding='utf-8') as f:
            f.write("has_changes=true\n")
    else:
        print("ℹ️ 所有幣種與交易所之階梯檔位與 MMR 數據無任何實質變動。")
        print("============================================================")
        with open(os.environ.get('GITHUB_OUTPUT', 'github_output.txt'), 'a', encoding='utf-8') as f:
            f.write("has_changes=false\n")
