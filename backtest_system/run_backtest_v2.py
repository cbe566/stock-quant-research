#!/usr/bin/env python3
"""
回測系統 v2 — 整合 Finnhub + FMP + 行業中性化 + 宏觀環境
"""

import sys, os, json
from datetime import datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import US_STOCKS, HK_STOCKS, FINNHUB_API_KEY, FMP_API_KEY, FRED_API_KEY
from data_engine import DataEngine
from strategies_v2 import StrategyV2
from finnhub_engine import FinnhubEngine
from backtester import Backtester
import pandas as pd
import numpy as np


def run_v2_backtest(tickers, market_name, start_date, end_date,
                    expected_return, timeframe_days, strategy):
    """v2 回測"""
    bt = Backtester()
    results = []

    print(f"\n{'='*80}")
    print(f"v2 回測：{market_name} | {start_date} → {end_date} | 目標:{expected_return}% / {timeframe_days}天")
    print(f"{'='*80}\n")

    for i, ticker in enumerate(tickers):
        print(f"[{i+1}/{len(tickers)}] {ticker}...", end=" ")
        try:
            analysis = strategy.comprehensive_analysis_v2(ticker, start_date, end_date)
            pred = analysis.get("prediction", {})
            direction = pred.get("direction", "觀望")
            score = pred.get("total_score", 50)

            if "看漲" in direction and "觀望" not in direction:
                result = bt.backtest_prediction(
                    ticker=ticker, prediction_date=start_date,
                    direction="漲", expected_return_pct=expected_return,
                    timeframe_days=timeframe_days, actual_end_date=end_date,
                )
                result["pred_direction"] = direction
                result["pred_score"] = score
                result["sector"] = analysis.get("sector", "")
                result["sector_type"] = analysis.get("sector_type", "")
                result["quality_score"] = analysis.get("quality_score", 0)
                result["value_score"] = analysis.get("value_score", 0)
                result["momentum_score"] = analysis.get("momentum_score", 0)
                result["analyst_score"] = analysis.get("analyst_score", 0)
                result["dcf_score"] = analysis.get("dcf_score", 0)
                result["tech_score"] = analysis.get("tech_score", 0)
                result["f_score"] = analysis.get("f_score", 0)
                result["bull_reasons"] = pred.get("bull_reasons", [])
                result["bear_reasons"] = pred.get("bear_reasons", [])
                results.append(result)
                status = "✓" if result.get("success") else "✗"
                print(f"{direction}(分:{score:.0f}) → {status} 實際:{result.get('final_return_pct',0):.1f}%")

            elif "看跌" in direction and "觀望" not in direction:
                result = bt.backtest_prediction(
                    ticker=ticker, prediction_date=start_date,
                    direction="跌", expected_return_pct=expected_return,
                    timeframe_days=timeframe_days, actual_end_date=end_date,
                )
                result["pred_direction"] = direction
                result["pred_score"] = score
                result["sector"] = analysis.get("sector", "")
                result["sector_type"] = analysis.get("sector_type", "")
                results.append(result)
                status = "✓" if result.get("success") else "✗"
                print(f"{direction}(分:{score:.0f}) → {status} 實際:{result.get('final_return_pct',0):.1f}%")

            else:
                print(f"{direction}(分:{score:.0f})")

        except Exception as e:
            print(f"錯誤: {str(e)[:60]}")
            continue

    return results


def calc_summary(results, label):
    """計算績效"""
    if not results:
        return {"label": label, "total": 0}

    df = pd.DataFrame(results)
    total = len(df)
    hits = int(df['success'].sum()) if 'success' in df.columns else 0
    wr = round(hits / total * 100, 1) if total > 0 else 0

    summary = {
        "label": label,
        "total": total,
        "wins": hits,
        "losses": total - hits,
        "win_rate": wr,
        "avg_return": round(df['final_return_pct'].mean(), 2) if 'final_return_pct' in df.columns else 0,
        "avg_max_fav": round(df['max_favorable_pct'].mean(), 2) if 'max_favorable_pct' in df.columns else 0,
        "avg_max_dd": round(df['max_drawdown_pct'].mean(), 2) if 'max_drawdown_pct' in df.columns else 0,
    }

    if 'final_return_pct' in df.columns and len(df) > 0:
        best_idx = df['final_return_pct'].idxmax()
        worst_idx = df['final_return_pct'].idxmin()
        summary["best"] = f"{df.loc[best_idx, 'ticker']} (+{df.loc[best_idx, 'final_return_pct']}%)"
        summary["worst"] = f"{df.loc[worst_idx, 'ticker']} ({df.loc[worst_idx, 'final_return_pct']}%)"

    return summary


def print_comparison(v1_summary, v2_summary):
    """v1 vs v2 比較"""
    print(f"\n{'█'*80}")
    print(f"{'█':^1}{'v1 vs v2 模型比較':^78}{'█':^1}")
    print(f"{'█'*80}\n")

    print(f"{'指標':<20} {'v1':>15} {'v2':>15} {'改善':>15}")
    print(f"{'─'*65}")

    for key, label in [("win_rate", "勝率(%)"), ("avg_return", "平均回報(%)"),
                        ("avg_max_dd", "平均最大回撤(%)")]:
        v1_val = v1_summary.get(key, 0)
        v2_val = v2_summary.get(key, 0)
        diff = v2_val - v1_val
        arrow = "↑" if diff > 0 else "↓" if diff < 0 else "→"
        print(f"{label:<20} {v1_val:>15.1f} {v2_val:>15.1f} {arrow}{abs(diff):>13.1f}")

    print(f"\n  v1 最佳：{v1_summary.get('best','N/A')}")
    print(f"  v2 最佳：{v2_summary.get('best','N/A')}")
    print(f"  v1 最差：{v1_summary.get('worst','N/A')}")
    print(f"  v2 最差：{v2_summary.get('worst','N/A')}")


# ==================== 主程式 ====================
if __name__ == "__main__":
    print("=" * 80)
    print("全方位量化回測系統 v2.0 — 優化版")
    print(f"啟動時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    # 初始化引擎
    de = DataEngine()
    fh = FinnhubEngine(FINNHUB_API_KEY)
    strategy = StrategyV2(de, finnhub=fh, fmp_key=FMP_API_KEY, fred_key=FRED_API_KEY)

    # 宏觀環境
    macro = strategy.get_macro_regime()
    weights = strategy.get_dynamic_weights(macro)
    print(f"\n宏觀環境：{macro['state'].upper()} | VIX:{macro.get('vix')} | 殖利率曲線:{macro.get('yield_curve')}")
    print(f"動態權重：質量{weights['quality']:.0%} 價值{weights['value']:.0%} "
          f"動量{weights['momentum']:.0%} 分析師{weights['analyst']:.0%} "
          f"內部人{weights['insider']:.0%} DCF{weights['dcf']:.0%}")

    # ============ 美股測試 ============
    all_us_results = []

    # 期間1：2024 Q1 → Q3
    r1 = run_v2_backtest(US_STOCKS, "美股-期間1", "2024-01-15", "2024-07-15", 10, 180, strategy)
    all_us_results.extend(r1)

    # 期間2：2024 Q3 → Q4
    r2 = run_v2_backtest(US_STOCKS, "美股-期間2", "2024-07-15", "2024-10-15", 8, 90, strategy)
    all_us_results.extend(r2)

    # 期間3：2024 Q4 → 2025 Q1
    r3 = run_v2_backtest(US_STOCKS, "美股-期間3", "2024-10-15", "2025-01-15", 8, 90, strategy)
    all_us_results.extend(r3)

    # ============ 港股測試 ============
    hk_results = run_v2_backtest(HK_STOCKS, "港股", "2024-01-15", "2024-07-15", 10, 180, strategy)

    # ============ 績效統計 ============
    us_summary = calc_summary(all_us_results, "美股v2")
    hk_summary_v2 = calc_summary(hk_results, "港股v2")

    print(f"\n{'█'*80}")
    print(f"{'█':^1}{'v2 回測結果':^78}{'█':^1}")
    print(f"{'█'*80}")

    print(f"\n📊 美股 v2")
    print(f"  總預測：{us_summary['total']} | 成功：{us_summary['wins']} | 失敗：{us_summary['losses']}")
    print(f"  勝率：{us_summary['win_rate']}%")
    print(f"  平均回報：{us_summary['avg_return']}%")
    print(f"  平均最大有利：{us_summary['avg_max_fav']}%")
    print(f"  平均最大回撤：{us_summary['avg_max_dd']}%")
    if us_summary.get('best'):
        print(f"  最佳：{us_summary['best']}")
    if us_summary.get('worst'):
        print(f"  最差：{us_summary['worst']}")

    print(f"\n📊 港股 v2")
    print(f"  總預測：{hk_summary_v2['total']} | 成功：{hk_summary_v2['wins']} | 失敗：{hk_summary_v2['losses']}")
    print(f"  勝率：{hk_summary_v2['win_rate']}%")
    print(f"  平均回報：{hk_summary_v2['avg_return']}%")

    # v1 數據（從之前回測結果）
    v1_summary = {"win_rate": 57.8, "avg_return": 11.45, "avg_max_dd": -15.57,
                  "best": "NVDA (+129.25%)", "worst": "AMD (-25.89%)"}
    print_comparison(v1_summary, us_summary)

    # 保存結果
    output = {
        "version": "v2",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "macro_regime": macro,
        "weights": weights,
        "us_summary": us_summary,
        "hk_summary": hk_summary_v2,
        "v1_comparison": v1_summary,
        "us_trades": [{k: v for k, v in r.items() if not isinstance(v, dict) or k in ["analysis"]}
                     for r in all_us_results],
    }

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backtest_v2_results.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n結果已保存：{output_path}")

    # 逐筆明細
    print(f"\n{'─'*80}")
    print(f"{'逐筆交易明細':^76}")
    print(f"{'─'*80}")

    for r in all_us_results:
        status = "✓" if r.get('success') else "✗"
        t = r.get('ticker', '?')
        d = r.get('pred_direction', '?')
        sc = r.get('pred_score', 0)
        f = r.get('final_return_pct', 0)
        sect = r.get('sector_type', '')

        print(f"\n  {status} {t:8s} [{sect:10s}] | 分數:{sc:5.1f} | 預測:{d:10s} | 實際:{f:>7.1f}%")

        # 分維度分數
        print(f"    質量:{r.get('quality_score',0):4.0f} 價值:{r.get('value_score',0):4.0f} "
              f"動量:{r.get('momentum_score',0):4.0f} 分析師:{r.get('analyst_score',0):4.0f} "
              f"DCF:{r.get('dcf_score',0):4.0f} 技術:{r.get('tech_score',0):4.0f} "
              f"F:{r.get('f_score',0)}/9")

        # 原因
        for reason in r.get('bull_reasons', [])[:2]:
            print(f"    ↑ {reason}")
        for reason in r.get('bear_reasons', [])[:2]:
            print(f"    ↓ {reason}")

    print(f"\n{'='*80}")
    print(f"完成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*80}")
