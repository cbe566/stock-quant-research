#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全方位回測系統 — 主執行腳本
測試市場：美股 + 港股
策略：多因子 + 技術面 + 動量 + 均值回歸
驗證：預測 vs 實際，逐筆分析成敗原因
"""

import sys
import os
import json
from datetime import datetime

# 確保能找到模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import US_STOCKS, HK_STOCKS, BENCHMARKS
from data_engine import DataEngine
from strategies import StrategyEngine
from backtester import Backtester

import pandas as pd
import numpy as np


def run_us_backtest():
    """美股全方位回測"""
    bt = Backtester()

    print("\n" + "█" * 80)
    print("█" + " " * 30 + "美 股 回 測" + " " * 31 + "█")
    print("█" * 80)

    # ============ 測試期間1：2024年Q1 → Q3（6個月持有）============
    print("\n\n▶ 測試期間 1：2024年Q1起，持有6個月")
    results_1 = bt.batch_backtest(
        tickers=US_STOCKS,
        start_date="2024-01-15",
        end_date="2024-07-15",
        expected_return=10,
        timeframe_days=180,
    )

    # ============ 測試期間2：2024年Q3 → Q4（3個月持有）============
    print("\n\n▶ 測試期間 2：2024年Q3起，持有3個月")
    results_2 = bt.batch_backtest(
        tickers=US_STOCKS,
        start_date="2024-07-15",
        end_date="2024-10-15",
        expected_return=8,
        timeframe_days=90,
    )

    # ============ 測試期間3：2024年Q4 → 2025年Q1（3個月持有）============
    print("\n\n▶ 測試期間 3：2024年Q4起，持有3個月")
    results_3 = bt.batch_backtest(
        tickers=US_STOCKS,
        start_date="2024-10-15",
        end_date="2025-01-15",
        expected_return=8,
        timeframe_days=90,
    )

    all_results = results_1 + results_2 + results_3
    summary = bt.performance_summary(all_results)
    bt.print_report(all_results, summary)

    return all_results, summary


def run_hk_backtest():
    """港股全方位回測"""
    bt = Backtester()

    print("\n\n" + "█" * 80)
    print("█" + " " * 30 + "港 股 回 測" + " " * 31 + "█")
    print("█" * 80)

    # 港股測試
    print("\n\n▶ 測試期間：2024年Q1起，持有6個月")
    results = bt.batch_backtest(
        tickers=HK_STOCKS,
        start_date="2024-01-15",
        end_date="2024-07-15",
        expected_return=10,
        timeframe_days=180,
    )

    summary = bt.performance_summary(results)
    bt.print_report(results, summary)

    return results, summary


def run_momentum_backtest():
    """動量策略專項回測"""
    de = DataEngine()
    se = StrategyEngine(de)

    print("\n\n" + "█" * 80)
    print("█" + " " * 26 + "動 量 策 略 回 測" + " " * 27 + "█")
    print("█" * 80)

    # 計算截面動量排名
    print("\n計算美股動量排名（2024年6月基準）...")
    ranking = se.momentum_ranking(US_STOCKS, "2023-01-01", "2024-06-30")

    if not ranking.empty:
        print("\n📊 動量排名（前10）：")
        print(ranking[['rank', 'ticker', 'raw_momentum', 'risk_adj_momentum', 'quintile']].head(10).to_string(index=False))

        print("\n📊 動量排名（後5 — 最弱）：")
        print(ranking[['rank', 'ticker', 'raw_momentum', 'risk_adj_momentum', 'quintile']].tail(5).to_string(index=False))

        # 驗證：買入前5強動量股，持有3個月
        top5 = ranking.head(5)['ticker'].tolist()
        bottom5 = ranking.tail(5)['ticker'].tolist()

        print(f"\n\n▶ 回測：買入動量前5（{', '.join(top5)}），持有3個月")
        bt = Backtester()
        results_top = bt.batch_backtest(
            tickers=top5,
            start_date="2024-07-01",
            end_date="2024-10-01",
            expected_return=5,
            timeframe_days=90,
        )

        print(f"\n▶ 對照：動量最弱5支（{', '.join(bottom5)}）同期表現")
        results_bottom = bt.batch_backtest(
            tickers=bottom5,
            start_date="2024-07-01",
            end_date="2024-10-01",
            expected_return=5,
            timeframe_days=90,
        )

        # 比較
        if results_top and results_bottom:
            top_avg = np.mean([r['final_return_pct'] for r in results_top if 'final_return_pct' in r])
            bottom_avg = np.mean([r['final_return_pct'] for r in results_bottom if 'final_return_pct' in r])
            print(f"\n📈 動量前5平均回報：{top_avg:.1f}%")
            print(f"📉 動量後5平均回報：{bottom_avg:.1f}%")
            print(f"📊 動量因子收益差：{top_avg - bottom_avg:.1f}%")

    return ranking


def save_results(us_results, us_summary, hk_results, hk_summary):
    """保存回測結果"""
    output_dir = os.path.dirname(os.path.abspath(__file__))

    # 保存 JSON
    report = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "us_market": {
            "summary": us_summary,
            "total_trades": len(us_results),
            "trades": [{k: v for k, v in r.items()
                       if k not in ['analysis']}
                      for r in us_results],
        },
        "hk_market": {
            "summary": hk_summary,
            "total_trades": len(hk_results),
            "trades": [{k: v for k, v in r.items()
                       if k not in ['analysis']}
                      for r in hk_results],
        },
    }

    json_path = os.path.join(output_dir, "backtest_results.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n結果已保存至：{json_path}")

    # 保存 CSV
    if us_results:
        df_us = pd.DataFrame([{k: v for k, v in r.items()
                               if not isinstance(v, (dict, list))}
                              for r in us_results])
        csv_path = os.path.join(output_dir, "backtest_us_trades.csv")
        df_us.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"美股交易明細：{csv_path}")

    if hk_results:
        df_hk = pd.DataFrame([{k: v for k, v in r.items()
                               if not isinstance(v, (dict, list))}
                              for r in hk_results])
        csv_path = os.path.join(output_dir, "backtest_hk_trades.csv")
        df_hk.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"港股交易明細：{csv_path}")


# ==================== 主程序 ====================
if __name__ == "__main__":
    print("=" * 80)
    print("全方位量化回測系統 v1.0")
    print(f"啟動時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    # 1. 美股回測
    us_results, us_summary = run_us_backtest()

    # 2. 港股回測
    hk_results, hk_summary = run_hk_backtest()

    # 3. 動量策略專項
    momentum_ranking = run_momentum_backtest()

    # 4. 保存結果
    save_results(us_results, us_summary, hk_results, hk_summary)

    # 5. 最終總結
    print("\n" + "█" * 80)
    print("█" + " " * 28 + "回 測 總 結" + " " * 29 + "█")
    print("█" * 80)

    print(f"\n  美股勝率：{us_summary.get('win_rate', 0)}%（{us_summary.get('successful', 0)}/{us_summary.get('total_predictions', 0)}）")
    print(f"  港股勝率：{hk_summary.get('win_rate', 0)}%（{hk_summary.get('successful', 0)}/{hk_summary.get('total_predictions', 0)}）")

    total_trades = us_summary.get('total_predictions', 0) + hk_summary.get('total_predictions', 0)
    total_wins = us_summary.get('successful', 0) + hk_summary.get('successful', 0)
    overall_wr = round(total_wins / total_trades * 100, 1) if total_trades > 0 else 0
    print(f"  綜合勝率：{overall_wr}%（{total_wins}/{total_trades}）")
    print(f"  美股平均回報：{us_summary.get('avg_final_return', 0)}%")
    print(f"  港股平均回報：{hk_summary.get('avg_final_return', 0)}%")

    print(f"\n完成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
