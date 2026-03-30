"""
合併 4 個市場的篩選結果 → 生成完整報告 → 上傳到 D1
用於 GitHub Actions matrix job 的最後合併步驟
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

def merge_market_results():
    """合併各市場 JSON 結果為完整報告"""
    artifacts_dir = Path("market_results")
    if not artifacts_dir.exists():
        print("❌ market_results/ 目錄不存在")
        sys.exit(1)

    # 讀取所有市場結果
    all_markets = {}
    all_errors = []
    date_str = None

    for json_file in sorted(artifacts_dir.glob("screening_*.json")):
        print(f"  📄 讀取 {json_file.name}")
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        market = data.get("market", "unknown")
        all_markets[market] = data.get("results", [])
        all_errors.extend(data.get("errors", []))
        if not date_str:
            date_str = data.get("date")

    if not all_markets:
        print("❌ 沒有找到任何市場結果")
        sys.exit(1)

    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    print(f"\n✅ 合併了 {len(all_markets)} 個市場的結果")
    for market, results in all_markets.items():
        print(f"   {market}: {len(results)} 支股票")

    # 儲存合併後的完整 JSON
    report_dir = Path("daily_reports")
    report_dir.mkdir(exist_ok=True)

    date_compact = date_str.replace("-", "")
    json_path = report_dir / f"screening_data_{date_compact}.json"
    merged_data = {
        "date": date_str,
        "generated_at": datetime.now().isoformat(),
        "markets": all_markets,
        "errors": all_errors,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(merged_data, f, ensure_ascii=False, indent=2, default=str)
    print(f"📊 合併數據已存到 {json_path}")

    # 生成 Markdown 報告（用 DailyScreener 的報告邏輯）
    try:
        sys.path.insert(0, os.path.dirname(__file__))
        from daily_screening import DailyScreener
        screener = DailyScreener()
        screener.results = all_markets
        screener.errors = all_errors
        md_path, report = screener.save_report()
        print(f"📄 Markdown 報告已生成：{md_path}")
    except Exception as e:
        print(f"⚠️ 報告生成失敗：{e}")
        # 至少保存 JSON 就好
        import traceback
        traceback.print_exc()

    return str(json_path)


def upload_to_d1(json_path=None):
    """上傳結果到 Cloudflare D1"""
    try:
        # 嘗試用現有的 upload_to_d1.py
        sys.path.insert(0, os.path.dirname(__file__))
        from upload_to_d1 import upload_screening_data
        success = upload_screening_data(json_path)
        if success:
            print("✅ D1 上傳成功")
        else:
            print("⚠️ D1 上傳回傳失敗")
    except Exception as e:
        print(f"⚠️ D1 上傳失敗（不影響報告）：{e}")


if __name__ == "__main__":
    json_path = merge_market_results()
    upload_to_d1(json_path)
