#!/usr/bin/env python3
"""
回測系統 v5 — 最佳組合版
策略：v2選股能力（67.9%勝率）+ 止盈提前出場 + 寬鬆止損 + 大盤過濾
核心邏輯：
- 用v2+盈餘修正的選股（已驗證有效）
- 止盈目標到達 → 提前鎖定利潤（v4證明止盈交易100%勝率，均報+24.7%）
- 止損放寬到12-15%（v4證明5%太窄，正常波動就被甩出）
- 大盤下跌時不開多單
"""

import sys, os, json
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import US_STOCKS, HK_STOCKS, FINNHUB_API_KEY, FMP_API_KEY, FRED_API_KEY
from data_engine import DataEngine
from strategies_v3 import StrategyV3
from finnhub_engine import FinnhubEngine
import pandas as pd
import numpy as np


class BacktestV5:
    def __init__(self, strategy):
        self.strategy = strategy
        self.de = strategy.de

    def get_spy_trend(self, date_str):
        try:
            start = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=60)).strftime("%Y-%m-%d")
            spy = self.de.get_prices("SPY", start, date_str)
            if spy.empty or len(spy) < 25:
                return "unknown"
            close = spy['Close']
            sma20 = float(close.rolling(20).mean().iloc[-1])
            sma50 = float(close.rolling(50).mean().iloc[-1]) if len(close) >= 50 else sma20
            current = float(close.iloc[-1])
            if current > sma20 and sma20 > sma50:
                return "strong_up"
            elif current > sma20:
                return "up"
            elif current < sma20 and current < sma50:
                return "down"
            else:
                return "neutral"
        except:
            return "unknown"

    def hybrid_exit(self, close_series, entry_price, direction,
                    take_profit=15, hard_stop=12, max_days=None):
        """
        混合出場：止盈主動出場 + 寬鬆硬止損 + 持有到期
        - 漲到止盈目標 → 立即出場鎖定利潤
        - 跌超過硬止損 → 認錯出場
        - 沒觸發任何一個 → 持有到期（像v2一樣）
        """
        if direction == "long":
            upper = entry_price * (1 + take_profit / 100)
            lower = entry_price * (1 - hard_stop / 100)
        else:  # short
            upper = entry_price * (1 - take_profit / 100)  # 做空的止盈是跌
            lower = entry_price * (1 + hard_stop / 100)    # 做空的止損是漲

        for i, (date, price) in enumerate(close_series.items()):
            pv = float(price)

            if direction == "long":
                if pv >= upper:
                    ret = (pv / entry_price - 1) * 100
                    return {"exit_reason": "止盈", "return_pct": round(ret, 2),
                            "holding_days": i, "exit_price": round(pv, 2)}
                if pv <= lower:
                    ret = (pv / entry_price - 1) * 100
                    return {"exit_reason": "止損", "return_pct": round(ret, 2),
                            "holding_days": i, "exit_price": round(pv, 2)}
            else:
                if pv <= upper:  # 做空：價格跌到目標
                    ret = (entry_price / pv - 1) * 100
                    return {"exit_reason": "止盈", "return_pct": round(ret, 2),
                            "holding_days": i, "exit_price": round(pv, 2)}
                if pv >= lower:  # 做空：價格漲超止損
                    ret = (entry_price / pv - 1) * 100
                    return {"exit_reason": "止損", "return_pct": round(ret, 2),
                            "holding_days": i, "exit_price": round(pv, 2)}

            if max_days and i >= max_days:
                break

        # 持有到期
        last = float(close_series.iloc[-1])
        if direction == "long":
            ret = (last / entry_price - 1) * 100
        else:
            ret = (entry_price / last - 1) * 100
        return {"exit_reason": "持有到期", "return_pct": round(ret, 2),
                "holding_days": len(close_series) - 1, "exit_price": round(last, 2)}

    def run_single(self, ticker, start, end, analysis, spy_trend, tp=15, stop=12):
        pred = analysis.get("prediction", {})
        direction = pred.get("direction", "觀望")
        score = pred.get("total_score", 50)

        if "觀望" in direction or "中性" in direction:
            return None

        # 大盤過濾：下跌時不做多
        if "看漲" in direction and spy_trend == "down":
            return None

        prices = self.de.get_prices(ticker, start, end)
        if prices.empty:
            return None

        close = prices['Close']
        entry_price = float(close.iloc[0])

        if "看漲" in direction:
            exit_r = self.hybrid_exit(close, entry_price, "long", tp, stop)
        elif "看跌" in direction:
            exit_r = self.hybrid_exit(close, entry_price, "short", tp, stop)
        else:
            return None

        returns = ((close / entry_price) - 1) * 100
        max_fav = float(returns.max()) if "看漲" in direction else float(-returns.min())
        max_dd = float(returns.min()) if "看漲" in direction else float(-returns.max())

        return {
            "ticker": ticker,
            "pred_direction": direction,
            "pred_score": round(score, 1),
            "spy_trend": spy_trend,
            "entry_price": round(entry_price, 2),
            "exit_price": exit_r.get("exit_price", 0),
            "exit_reason": exit_r.get("exit_reason", ""),
            "return_pct": exit_r.get("return_pct", 0),
            "holding_days": exit_r.get("holding_days", 0),
            "max_favorable": round(max_fav, 2),
            "max_drawdown": round(max_dd, 2),
            "success": exit_r.get("return_pct", 0) > 0,
            "sector": analysis.get("sector", ""),
            "revision_score": analysis.get("revision_score", 50),
            "f_score": analysis.get("f_score", 0),
        }


def run_batch(tickers, label, start, end, tp, stop, bt, strategy):
    results = []
    spy_trend = bt.get_spy_trend(start)

    print(f"\n{'='*80}")
    print(f"v5：{label} | {start}→{end} | SPY:{spy_trend} | 止盈:{tp}% 止損:{stop}%")
    print(f"{'='*80}\n")

    for i, t in enumerate(tickers):
        print(f"[{i+1}/{len(tickers)}] {t}...", end=" ")
        try:
            a = strategy.comprehensive_analysis_v3(t, start, end)
            r = bt.run_single(t, start, end, a, spy_trend, tp, stop)
            if r:
                results.append(r)
                s = "✓" if r["success"] else "✗"
                print(f"{s} {r['pred_direction']} (分:{r['pred_score']:.0f}) "
                      f"→ {r['return_pct']:+.1f}% ({r['exit_reason']}, {r['holding_days']}天)")
            else:
                d = a.get("prediction", {}).get("direction", "觀望")
                sc = a.get("prediction", {}).get("total_score", 50)
                if spy_trend == "down" and "看漲" in d:
                    print(f"{d}(分:{sc:.0f}) — 大盤過濾")
                else:
                    print(f"{d}(分:{sc:.0f}) — 跳過")
        except Exception as e:
            print(f"錯誤:{str(e)[:40]}")

    return results


def calc_stats(results, label):
    if not results:
        return {"label": label, "total": 0}
    df = pd.DataFrame(results)
    t = len(df)
    w = int(df['success'].sum())
    wr = round(w / t * 100, 1)

    s = {"label": label, "total": t, "wins": w, "losses": t - w, "win_rate": wr,
         "avg_return": round(df['return_pct'].mean(), 2),
         "total_return": round(df['return_pct'].sum(), 2),
         "avg_holding": round(df['holding_days'].mean(), 0),
         "profit_factor": 0}

    # 盈虧比
    wins_df = df[df['return_pct'] > 0]
    losses_df = df[df['return_pct'] <= 0]
    avg_win = wins_df['return_pct'].mean() if len(wins_df) > 0 else 0
    avg_loss = abs(losses_df['return_pct'].mean()) if len(losses_df) > 0 else 1
    s["avg_win"] = round(avg_win, 2)
    s["avg_loss"] = round(avg_loss, 2)
    s["win_loss_ratio"] = round(avg_win / avg_loss, 2) if avg_loss > 0 else 999
    s["profit_factor"] = round(wins_df['return_pct'].sum() / abs(losses_df['return_pct'].sum()), 2) if len(losses_df) > 0 and losses_df['return_pct'].sum() != 0 else 999
    s["expectancy"] = round(wr/100 * avg_win - (1-wr/100) * avg_loss, 2)

    # 出場統計
    for reason in df['exit_reason'].unique():
        sub = df[df['exit_reason'] == reason]
        s[f"exit_{reason}"] = {
            "count": len(sub), "wr": round(sub['success'].mean()*100, 1),
            "avg": round(sub['return_pct'].mean(), 2)
        }

    if t > 0:
        s["best"] = f"{df.loc[df['return_pct'].idxmax(), 'ticker']} ({df['return_pct'].max():+.1f}%)"
        s["worst"] = f"{df.loc[df['return_pct'].idxmin(), 'ticker']} ({df['return_pct'].min():+.1f}%)"

    return s


if __name__ == "__main__":
    print("=" * 80)
    print("v5.0 — 最佳組合：v2選股 + 止盈鎖利 + 寬鬆止損 + 大盤過濾")
    print(f"時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    de = DataEngine()
    fh = FinnhubEngine(FINNHUB_API_KEY)
    strat = StrategyV3(de, finnhub=fh, fmp_key=FMP_API_KEY, fred_key=FRED_API_KEY)
    bt = BacktestV5(strat)

    all_us = []

    # 期間1：6個月，止盈15%/止損12%
    r1 = run_batch(US_STOCKS, "美股P1(6M)", "2024-01-15", "2024-07-15", 15, 12, bt, strat)
    all_us.extend(r1)

    # 期間2：3個月，止盈12%/止損10%
    r2 = run_batch(US_STOCKS, "美股P2(3M)", "2024-07-15", "2024-10-15", 12, 10, bt, strat)
    all_us.extend(r2)

    # 期間3：3個月，止盈12%/止損10%
    r3 = run_batch(US_STOCKS, "美股P3(3M)", "2024-10-15", "2025-01-15", 12, 10, bt, strat)
    all_us.extend(r3)

    # 港股
    hk = run_batch(HK_STOCKS, "港股(6M)", "2024-01-15", "2024-07-15", 15, 12, bt, strat)

    us_s = calc_stats(all_us, "美股v5")
    hk_s = calc_stats(hk, "港股v5")

    print(f"\n{'█'*80}")
    print(f"{'完整演進 v1→v2→v3→v4→v5':^74}")
    print(f"{'█'*80}\n")

    print(f"{'指標':<18} {'v1':>8} {'v2':>8} {'v3':>8} {'v4':>8} {'v5':>8}")
    print(f"{'─'*58}")
    print(f"{'勝率(%)':<18} {'57.8':>8} {'67.9':>8} {'47.2':>8} {'41.7':>8} {us_s['win_rate']:>8.1f}")
    print(f"{'平均回報(%)':<18} {'11.4':>8} {'9.6':>8} {'2.3':>8} {'3.1':>8} {us_s['avg_return']:>8.2f}")
    print(f"{'預測數量':<18} {'64':>8} {'81':>8} {'36':>8} {'36':>8} {us_s['total']:>8}")
    print(f"{'盈虧比':<18} {'—':>8} {'—':>8} {'—':>8} {'—':>8} {us_s['win_loss_ratio']:>8.2f}")
    print(f"{'利潤因子':<18} {'—':>8} {'—':>8} {'—':>8} {'—':>8} {us_s['profit_factor']:>8.2f}")
    print(f"{'期望值(%)':<18} {'—':>8} {'—':>8} {'—':>8} {'—':>8} {us_s['expectancy']:>8.2f}")

    print(f"\n📊 出場原因：")
    for k, v in us_s.items():
        if k.startswith("exit_") and isinstance(v, dict):
            name = k.replace("exit_", "")
            print(f"  {name}：{v['count']}筆 | 勝率:{v['wr']}% | 均報:{v['avg']}%")

    print(f"\n  最佳：{us_s.get('best','—')}")
    print(f"  最差：{us_s.get('worst','—')}")
    print(f"  平均持有：{us_s.get('avg_holding',0):.0f}天")

    print(f"\n📊 港股v5：勝率{hk_s.get('win_rate',0)}% | 均報:{hk_s.get('avg_return',0)}%")

    # 逐筆
    print(f"\n{'─'*80}")
    for r in all_us:
        s = "✓" if r['success'] else "✗"
        print(f"  {s} {r['ticker']:6s} | {r['pred_direction']:8s} 分:{r['pred_score']:5.1f} "
              f"| {r['return_pct']:>+7.1f}% | {r['exit_reason']:6s} {r['holding_days']:>3}天 "
              f"| 最大有利:{r['max_favorable']:>5.1f}% 回撤:{r['max_drawdown']:>6.1f}%")

    # 保存
    out = {"version": "v5", "generated_at": datetime.now().isoformat(),
           "us": us_s, "hk": hk_s,
           "evolution": {"v1": 57.8, "v2": 67.9, "v3": 47.2, "v4": 41.7, "v5": us_s['win_rate']}}
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backtest_v5_results.json")
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n{'='*80}")
    print(f"完成 | {path}")
    print(f"{'='*80}")
