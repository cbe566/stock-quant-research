"""
上傳篩選結果到 Cloudflare D1
============================
GitHub Actions 篩選完成後呼叫此腳本，將 JSON 結果推送到 Worker API
"""

import json
import sys
import os
import requests
from datetime import datetime
from pathlib import Path


# Worker API 端點
WORKER_URL = "https://stock-screening-api.stock-quant.workers.dev/api/upload"
API_KEY = os.environ.get("WORKER_API_KEY", "stock-screening-2026")


def upload_screening_data(json_path=None):
    """讀取篩選結果 JSON 並上傳到 D1"""

    # 自動找最新的 JSON 檔案
    if json_path is None:
        report_dir = Path(__file__).parent / "daily_reports"
        today = datetime.now().strftime("%Y%m%d")
        json_path = report_dir / f"screening_data_{today}.json"

        if not json_path.exists():
            # 找最新的
            json_files = sorted(report_dir.glob("screening_data_*.json"), reverse=True)
            if json_files:
                json_path = json_files[0]
            else:
                print("找不到篩選數據 JSON 檔案")
                return False

    print(f"讀取數據：{json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    date = data.get("date", datetime.now().strftime("%Y-%m-%d"))
    markets = data.get("markets", {})

    if not markets:
        print("數據為空，跳過上傳")
        return False

    # 統計
    total = sum(len(v) for v in markets.values())
    print(f"日期：{date} | 總股數：{total}")

    # 上傳到 Worker API
    payload = {
        "date": date,
        "markets": markets,
    }

    print(f"正在上傳到 D1...")
    try:
        resp = requests.post(
            WORKER_URL,
            json=payload,
            headers={
                "Authorization": f"Bearer {API_KEY}",
                "Content-Type": "application/json",
            },
            timeout=30,
        )

        if resp.status_code == 200:
            result = resp.json()
            print(f"上傳成功！插入 {result.get('inserted', 0)} 筆記錄")
            return True
        else:
            print(f"上傳失敗：HTTP {resp.status_code}")
            print(resp.text[:500])
            return False

    except Exception as e:
        print(f"上傳出錯：{e}")
        return False


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else None
    success = upload_screening_data(path)
    sys.exit(0 if success else 1)
