#!/usr/bin/env python3
"""
回測系統 v3 — 加入盈餘修正、Triple Barrier出場、信心過濾
目標：勝率 75%+，回報 12%+
"""

import sys, os, json
from datetime import datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import US_STOCKS, HK_STOCKS, FINNHUB_API_KEY, FMP_API_KEY, FRED_API_KEY
from data_engine import DataEngine
from strategies_v3 import StrategyV3
from finnhub_engine import FinnhubEngine
import pandas as pd
import numpy as np


def backtest_with_triple_barrier(de, strategy, ticker, pred_direction, pred_score,
                                  confidence, position_size, analysis_result,
                                  start_date, end_date, take_profit=15, stop_loss=8):
    """用 Triple Barrier 出場策略回測"""
    # 取得價格數據
    prices = de.get_prices(ticker, start_date, end_date)
    if prices.empty:
        return None

    close = prices['Close']
    entry_date = close.index[0]
    entry_price = float(close.iloc[0])

    # Triple Barrier 出場
    if "看漲" in pred_direction:
        exit_result = strategy.triple_barrier_exit(
            close, entry_price, take_profit_pct=take_profit,
            stop_loss_pct=stop_loss, max_holding_days=len(close)
        )
    elif "看跌" in pred_direction:
        # 反轉：止盈=價格跌到目標，止損=價格漲超閾值
        inverted = entry_price * 2 - close  # 鏡像價格
        exit_result = strategy.triple_barrier_exit(
            close, entry_price, take_profit_pct=take_profit,
            stop_loss_pct=stop_loss, max_holding_days=len(close)
        )
        # 看跌的回報取反
        exit_result["return_pct"] = -exit_result.get("return_pct", 0)
    else:
        return None

    # 計算最大有利和最大回撤
    returns = ((close / entry_price) - 1) * 100
    max_favorable = float(returns.max())
    max_adverse = float(returns.min())

    success = exit_result["return_pct"] > 0

    return {
        "ticker": ticker,
        "pred_direction": pred_direction,
        "pred_score": pred_score,
        "confidence": confidence,
        "position_size": position_size,
        "entry_price": entry_price,
        "exit_price": exit_result.get("exit_price", 0),
        "exit_reason": exit_result.get("exit_reason", ""),
        "exit_return_pct": exit_result.get("return_pct", 0),
        "holding_days": exit_result.get("holding_days", 0),
        "max_favorable_pct": round(max_favorable, 2),
        "max_drawdown_pct": round(max_adverse, 2),
        "success": success,
        "final_return_pct": exit_result.get("return_pct", 0),
        "sector": analysis_result.get("sector", ""),
        "sector_type": analysis_result.get("sector_type", ""),
        "quality_score": analysis_result.get("quality_score", 0),
        "value_score": analysis_result.get("value_score", 0),
        "momentum_score": analysis_result.get("momentum_score", 0),
        "analyst_score": analysis_result.get("analyst_score", 0),
        "dcf_score": analysis_result.get("dcf_score", 0),
        "tech_score": analysis_result.get("tech_score", 0),
        "revision_score": analysis_result.get("revision_score", 0),
        "f_score": analysis_result.get("f_score", 0),
        "bull_reasons": analysis_result.get("prediction", {}).get("bull_reasons", []),
        "bear_reasons": analysis_result.get("prediction", {}).get("bear_reasons", []),
    }


def run_v3_backtest(tickers, market_name, start_date, end_date,
                    take_profit, stop_loss, strategy):
    """v3 回測（含 Triple Barrier）"""
    results = []

    print(f"\n{'='*80}")
    print(f"v3 回測：{market_name} | {start_date} → {end_date}")
    print(f"Triple Barrier：止盈{take_profit}% / 止損{stop_loss}%")
    print(f"{'='*80}\n")

    for i, ticker in enumerate(tickers):
        print(f"[{i+1}/{len(tickers)}] {ticker}...", end=" ")
        try:
            analysis = strategy.comprehensive_analysis_v3(ticker, start_date, end_date)
            pred = analysis.get("prediction", {})
            direction = pred.get("direction", "觀望")
            score = pred.get("total_score", 50)
            confidence = pred.get("confidence", "C")
            position = pred.get("position_size", 0)

            # P0-3: 只有A級和B級信心才出手
            if confidence == "C" or "觀望" in direction or "中性" in direction:
                print(f"{direction} [{confidence}] (分:{score:.0f}) — 跳過")
                continue

            if "看漲" in direction or "看跌" in direction:
                result = backtest_with_triple_barrier(
                    strategy.de, strategy, ticker, direction, score,
                    confidence, position, analysis,
                    start_date, end_date, take_profit, stop_loss,
                )
                if result:
                    results.append(result)
                    status = "✓" if result["success"] else "✗"
                    exit_r = result["exit_reason"]
                    ret = result["exit_return_pct"]
                    days = result["holding_days"]
                    print(f"{status} {direction} [{confidence}] (分:{score:.0f}) "
                          f"→ {ret:+.1f}% ({exit_r}, {days}天)")
                else:
                    print(f"{direction} [{confidence}] — 無數據")
            else:
                print(f"{direction} [{confidence}] (分:{score:.0f})")

        except Exception as e:
            print(f"錯誤: {str(e)[:50]}")
            continue

    return results


def calc_summary(results, label):
    if not results:
        return {"label": label, "total": 0, "win_rate": 0, "avg_return": 0}

    df = pd.DataFrame(results)
    total = len(df)
    hits = int(df['success'].sum())
    wr = round(hits / total * 100, 1) if total > 0 else 0

    # 加權回報（信心加權）
    if 'position_size' in df.columns and 'exit_return_pct' in df.columns:
        weighted_ret = (df['exit_return_pct'] * df['position_size']).sum() / df['position_size'].sum()
    else:
        weighted_ret = df['final_return_pct'].mean() if 'final_return_pct' in df.columns else 0

    # 按出場原因統計
    exit_stats = {}
    if 'exit_reason' in df.columns:
        for reason in df['exit_reason'].unique():
            sub = df[df['exit_reason'] == reason]
            exit_stats[reason] = {
                "count": len(sub),
                "avg_return": round(sub['exit_return_pct'].mean(), 2),
                "win_rate": round(sub['success'].mean() * 100, 1),
            }

    # 按信心等級統計
    conf_stats = {}
    if 'confidence' in df.columns:
        for conf in ['A', 'B']:
            sub = df[df['confidence'] == conf]
            if len(sub) > 0:
                conf_stats[conf] = {
                    "count": len(sub),
                    "win_rate": round(sub['success'].mean() * 100, 1),
                    "avg_return": round(sub['exit_return_pct'].mean(), 2),
                }

    summary = {
        "label": label,
        "total": total,
        "wins": hits,
        "losses": total - hits,
        "win_rate": wr,
        "avg_return": round(df['exit_return_pct'].mean(), 2) if 'exit_return_pct' in df.columns else 0,
        "weighted_return": round(weighted_ret, 2),
        "avg_holding_days": round(df['holding_days'].mean(), 0) if 'holding_days' in df.columns else 0,
        "exit_stats": exit_stats,
        "confidence_stats": conf_stats,
    }

    if 'exit_return_pct' in df.columns and len(df) > 0:
        best_idx = df['exit_return_pct'].idxmax()
        worst_idx = df['exit_return_pct'].idxmin()
        summary["best"] = f"{df.loc[best_idx, 'ticker']} ({df.loc[best_idx, 'exit_return_pct']:+.1f}%)"
        summary["worst"] = f"{df.loc[worst_idx, 'ticker']} ({df.loc[worst_idx, 'exit_return_pct']:+.1f}%)"

    return summary


# ==================== 主程式 ====================
if __name__ == "__main__":
    print("=" * 80)
    print("全方位量化回測系統 v3.0 — 盈餘修正 + Triple Barrier + 信心過濾")
    print(f"啟動時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    de = DataEngine()
    fh = FinnhubEngine(FINNHUB_API_KEY)
    strategy = StrategyV3(de, finnhub=fh, fmp_key=FMP_API_KEY, fred_key=FRED_API_KEY)

    # 宏觀環境
    macro = strategy.get_macro_regime()
    print(f"\n宏觀環境：{macro['state'].upper()} | VIX:{macro.get('vix')}")

    all_us = []

    # 期間1：2024 Q1→Q3（止盈15%/止損8%）
    r1 = run_v3_backtest(US_STOCKS, "美股-期間1(6M)", "2024-01-15", "2024-07-15", 15, 8, strategy)
    all_us.extend(r1)

    # 期間2：2024 Q3→Q4（止盈12%/止損7%，短期）
    r2 = run_v3_backtest(US_STOCKS, "美股-期間2(3M)", "2024-07-15", "2024-10-15", 12, 7, strategy)
    all_us.extend(r2)

    # 期間3：2024 Q4→2025 Q1（止盈12%/止損7%）
    r3 = run_v3_backtest(US_STOCKS, "美股-期間3(3M)", "2024-10-15", "2025-01-15", 12, 7, strategy)
    all_us.extend(r3)

    # 港股
    hk = run_v3_backtest(HK_STOCKS, "港股(6M)", "2024-01-15", "2024-07-15", 15, 8, strategy)

    # ============ 績效統計 ============
    us_sum = calc_summary(all_us, "美股v3")
    hk_sum = calc_summary(hk, "港股v3")

    print(f"\n{'█'*80}")
    print(f"{'v1 vs v2 vs v3 完整比較':^74}")
    print(f"{'█'*80}")

    print(f"\n{'指標':<20} {'v1':>12} {'v2':>12} {'v3':>12} {'v2→v3':>12}")
    print(f"{'─'*68}")
    v1_wr, v2_wr, v3_wr = 57.8, 67.9, us_sum['win_rate']
    v1_ret, v2_ret, v3_ret = 11.45, 9.56, us_sum['avg_return']
    print(f"{'勝率(%)':<20} {v1_wr:>12.1f} {v2_wr:>12.1f} {v3_wr:>12.1f} {v3_wr-v2_wr:>+12.1f}")
    print(f"{'平均回報(%)':<20} {v1_ret:>12.2f} {v2_ret:>12.2f} {v3_ret:>12.2f} {v3_ret-v2_ret:>+12.2f}")
    print(f"{'預測數量':<20} {'64':>12} {'81':>12} {us_sum['total']:>12}")
    print(f"{'平均持有天數':<20} {'—':>12} {'—':>12} {us_sum['avg_holding_days']:>12.0f}")
    print(f"{'加權回報(%)':<20} {'—':>12} {'—':>12} {us_sum['weighted_return']:>12.2f}")

    # 出場原因統計
    print(f"\n📊 出場原因分析：")
    for reason, stats in us_sum.get("exit_stats", {}).items():
        print(f"  {reason}：{stats['count']}筆 | 勝率:{stats['win_rate']}% | 平均回報:{stats['avg_return']}%")

    # 信心等級統計
    print(f"\n📊 信心等級分析：")
    for conf, stats in us_sum.get("confidence_stats", {}).items():
        print(f"  {conf}級：{stats['count']}筆 | 勝率:{stats['win_rate']}% | 平均回報:{stats['avg_return']}%")

    if us_sum.get("best"):
        print(f"\n  最佳交易：{us_sum['best']}")
    if us_sum.get("worst"):
        print(f"  最差交易：{us_sum['worst']}")

    print(f"\n📊 港股 v3")
    print(f"  勝率：{hk_sum['win_rate']}%（{hk_sum['wins']}/{hk_sum['total']}）")
    print(f"  平均回報：{hk_sum['avg_return']}%")

    # 逐筆明細
    print(f"\n{'─'*80}")
    print(f"{'逐筆交易明細':^76}")
    print(f"{'─'*80}")

    for r in all_us:
        s = "✓" if r['success'] else "✗"
        print(f"\n  {s} {r['ticker']:8s} [{r.get('confidence','?')}] | "
              f"分:{r['pred_score']:5.1f} | {r['pred_direction']:8s} | "
              f"回報:{r['exit_return_pct']:>+7.1f}% | "
              f"{r['exit_reason']}({r['holding_days']}天)")
        print(f"    質:{r.get('quality_score',0):3.0f} 值:{r.get('value_score',0):3.0f} "
              f"量:{r.get('momentum_score',0):3.0f} 師:{r.get('analyst_score',0):3.0f} "
              f"DCF:{r.get('dcf_score',0):3.0f} 技:{r.get('tech_score',0):3.0f} "
              f"修:{r.get('revision_score',0):3.0f} F:{r.get('f_score',0)}/9")

    # 保存
    output = {
        "version": "v3", "generated_at": datetime.now().isoformat(),
        "improvements": ["P0-1:盈餘修正因子", "P0-2:Triple Barrier出場", "P0-3:信心過濾"],
        "us_summary": us_sum, "hk_summary": hk_sum,
        "comparison": {"v1_wr": 57.8, "v2_wr": 67.9, "v3_wr": us_sum['win_rate']},
    }
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backtest_v3_results.json")
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n{'='*80}")
    print(f"完成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"結果已保存：{path}")
    print(f"{'='*80}")
