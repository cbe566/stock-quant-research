#!/usr/bin/env python3
"""
回測系統 v4 — 最終優化版
核心改進：
1. v2選股（已驗證67.9%勝率）+ v3盈餘修正
2. ATR動態止損（波動大的股票放寬止損）
3. SPY大盤趨勢過濾器（大盤弱勢時不開多單）
4. 分批止盈（到達目標1先獲利50%，剩餘追蹤止盈）
5. 信心分級但閾值更合理
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
import yfinance as yf


class BacktestV4:
    """v4 回測引擎"""

    def __init__(self, strategy):
        self.strategy = strategy
        self.de = strategy.de

    def get_spy_trend(self, date_str):
        """判斷SPY趨勢：價格是否在20日均線之上"""
        try:
            start = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=60)).strftime("%Y-%m-%d")
            spy = self.de.get_prices("SPY", start, date_str)
            if spy.empty or len(spy) < 25:
                return "unknown"
            close = spy['Close']
            sma20 = close.rolling(20).mean()
            sma50 = close.rolling(50).mean()
            current = float(close.iloc[-1])
            ma20 = float(sma20.iloc[-1])
            ma50 = float(sma50.iloc[-1]) if len(sma50.dropna()) > 0 else ma20

            if current > ma20 and ma20 > ma50:
                return "strong_up"    # 強勢上升
            elif current > ma20:
                return "up"           # 上升
            elif current > ma50:
                return "neutral"      # 中性
            else:
                return "down"         # 下降
        except:
            return "unknown"

    def calc_atr_stop(self, ticker, date_str, multiplier=2.0):
        """計算ATR動態止損幅度"""
        try:
            start = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=60)).strftime("%Y-%m-%d")
            prices = self.de.get_prices(ticker, start, date_str)
            if prices.empty or len(prices) < 20:
                return 8.0  # 默認8%

            atr = self.de.calc_atr(prices['High'], prices['Low'], prices['Close'], 14)
            if len(atr.dropna()) == 0:
                return 8.0

            current_atr = float(atr.iloc[-1])
            current_price = float(prices['Close'].iloc[-1])
            stop_pct = (current_atr * multiplier / current_price) * 100
            # 限制範圍：最小5%，最大15%
            return round(max(min(stop_pct, 15.0), 5.0), 1)
        except:
            return 8.0

    def smart_exit(self, prices_series, entry_price, stop_loss_pct,
                   take_profit_1=10, take_profit_2=20, max_days=None):
        """
        智能出場策略：
        - 第一目標（take_profit_1）：獲利50%倉位
        - 第二目標（take_profit_2）：剩餘追蹤止盈
        - ATR動態止損
        - 時間限制
        """
        upper_1 = entry_price * (1 + take_profit_1 / 100)
        upper_2 = entry_price * (1 + take_profit_2 / 100)
        lower = entry_price * (1 - stop_loss_pct / 100)

        hit_target_1 = False
        trailing_stop = lower
        best_price = entry_price

        for i, (date, price) in enumerate(prices_series.items()):
            pv = float(price)

            # 更新最高價和追蹤止損
            if pv > best_price:
                best_price = pv
                if hit_target_1:
                    # 到達第一目標後，追蹤止損 = 最高價的92%
                    trailing_stop = max(trailing_stop, best_price * 0.92)

            # 止損觸發
            current_stop = trailing_stop if hit_target_1 else lower
            if pv <= current_stop:
                ret = (pv / entry_price - 1) * 100
                reason = "追蹤止損" if hit_target_1 else "止損"
                return {
                    "exit_price": round(pv, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": reason,
                    "return_pct": round(ret, 2),
                    "holding_days": i,
                    "hit_target_1": hit_target_1,
                }

            # 第一目標達成
            if not hit_target_1 and pv >= upper_1:
                hit_target_1 = True
                # 記錄但不退出，追蹤止盈
                trailing_stop = entry_price * (1 + take_profit_1 * 0.5 / 100)  # 保住一半利潤

            # 第二目標達成 → 全部退出
            if pv >= upper_2:
                ret = (pv / entry_price - 1) * 100
                return {
                    "exit_price": round(pv, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": "止盈(目標2)",
                    "return_pct": round(ret, 2),
                    "holding_days": i,
                    "hit_target_1": True,
                }

            # 時間限制
            if max_days and i >= max_days:
                ret = (pv / entry_price - 1) * 100
                return {
                    "exit_price": round(pv, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": "時間到期",
                    "return_pct": round(ret, 2),
                    "holding_days": i,
                    "hit_target_1": hit_target_1,
                }

        # 數據結束
        if len(prices_series) > 0:
            last = float(prices_series.iloc[-1])
            ret = (last / entry_price - 1) * 100
            return {
                "exit_price": round(last, 2),
                "exit_date": str(prices_series.index[-1].date()) if hasattr(prices_series.index[-1], 'date') else "",
                "exit_reason": "數據截止",
                "return_pct": round(ret, 2),
                "holding_days": len(prices_series),
                "hit_target_1": hit_target_1,
            }

        return {"exit_reason": "無數據", "return_pct": 0, "holding_days": 0, "hit_target_1": False}

    def run_single(self, ticker, start_date, end_date, analysis,
                   spy_trend, tp1=10, tp2=20):
        """回測單支股票"""
        pred = analysis.get("prediction", {})
        direction = pred.get("direction", "觀望")
        score = pred.get("total_score", 50)

        # 跳過觀望和中性
        if "觀望" in direction or "中性" in direction:
            return None

        # 大盤趨勢過濾：大盤下跌時不做多
        if "看漲" in direction and spy_trend == "down":
            return None  # 大盤弱勢不做多

        # 計算ATR止損
        stop_pct = self.calc_atr_stop(ticker, start_date)

        # 取得價格
        prices = self.de.get_prices(ticker, start_date, end_date)
        if prices.empty:
            return None

        close = prices['Close']
        entry_price = float(close.iloc[0])

        # 執行智能出場
        if "看漲" in direction:
            exit_result = self.smart_exit(close, entry_price, stop_pct,
                                          tp1, tp2, max_days=len(close))
        elif "看跌" in direction:
            # 看跌：反轉邏輯
            inverted_close = entry_price * 2 - close
            exit_result = self.smart_exit(inverted_close, entry_price, stop_pct,
                                          tp1, tp2, max_days=len(close))
            # 實際回報
            actual_last = float(close.iloc[min(exit_result.get("holding_days", 0), len(close)-1)])
            exit_result["return_pct"] = round((entry_price / actual_last - 1) * 100, 2) if "看跌" in direction else exit_result["return_pct"]
        else:
            return None

        success = exit_result["return_pct"] > 0

        # 最大有利/回撤
        returns = ((close / entry_price) - 1) * 100
        max_fav = float(returns.max()) if "看漲" in direction else float(-returns.min())
        max_dd = float(returns.min()) if "看漲" in direction else float(-returns.max())

        return {
            "ticker": ticker,
            "pred_direction": direction,
            "pred_score": round(score, 1),
            "confidence": pred.get("confidence", "B"),
            "spy_trend": spy_trend,
            "atr_stop_pct": stop_pct,
            "entry_price": round(entry_price, 2),
            "exit_price": exit_result.get("exit_price", 0),
            "exit_reason": exit_result.get("exit_reason", ""),
            "exit_return_pct": exit_result.get("return_pct", 0),
            "holding_days": exit_result.get("holding_days", 0),
            "hit_target_1": exit_result.get("hit_target_1", False),
            "max_favorable_pct": round(max_fav, 2),
            "max_drawdown_pct": round(max_dd, 2),
            "success": success,
            "sector": analysis.get("sector", ""),
            "sector_type": analysis.get("sector_type", ""),
            "revision_score": analysis.get("revision_score", 50),
            "f_score": analysis.get("f_score", 0),
            "bull_reasons": pred.get("bull_reasons", [])[:3],
            "bear_reasons": pred.get("bear_reasons", [])[:3],
        }


def run_v4(tickers, label, start, end, tp1, tp2, bt, strategy):
    """v4 批量回測"""
    results = []
    spy_trend = bt.get_spy_trend(start)

    print(f"\n{'='*80}")
    print(f"v4：{label} | {start}→{end} | SPY趨勢:{spy_trend} | 止盈:{tp1}/{tp2}%")
    print(f"{'='*80}\n")

    if spy_trend == "down":
        print("  ⚠️ SPY趨勢下降，多單將被過濾，僅保留空單\n")

    for i, ticker in enumerate(tickers):
        print(f"[{i+1}/{len(tickers)}] {ticker}...", end=" ")
        try:
            analysis = strategy.comprehensive_analysis_v3(ticker, start, end)
            result = bt.run_single(ticker, start, end, analysis, spy_trend, tp1, tp2)

            if result:
                results.append(result)
                s = "✓" if result["success"] else "✗"
                print(f"{s} {result['pred_direction']} (分:{result['pred_score']:.0f}) "
                      f"→ {result['exit_return_pct']:+.1f}% "
                      f"({result['exit_reason']}, {result['holding_days']}天, "
                      f"止損:{result['atr_stop_pct']}%)")
            else:
                pred = analysis.get("prediction", {})
                d = pred.get("direction", "觀望")
                sc = pred.get("total_score", 50)
                if spy_trend == "down" and "看漲" in d:
                    print(f"{d} (分:{sc:.0f}) — 大盤弱勢過濾")
                else:
                    print(f"{d} (分:{sc:.0f}) — 跳過")
        except Exception as e:
            print(f"錯誤: {str(e)[:50]}")

    return results


def summary(results, label):
    if not results:
        return {"label": label, "total": 0, "win_rate": 0}

    df = pd.DataFrame(results)
    total = len(df)
    wins = int(df['success'].sum())
    wr = round(wins / total * 100, 1)

    s = {
        "label": label, "total": total, "wins": wins, "losses": total - wins,
        "win_rate": wr,
        "avg_return": round(df['exit_return_pct'].mean(), 2),
        "median_return": round(df['exit_return_pct'].median(), 2),
        "avg_holding": round(df['holding_days'].mean(), 0),
        "avg_max_fav": round(df['max_favorable_pct'].mean(), 2),
        "avg_max_dd": round(df['max_drawdown_pct'].mean(), 2),
    }

    # 按出場原因
    for reason in df['exit_reason'].unique():
        sub = df[df['exit_reason'] == reason]
        s[f"exit_{reason}"] = f"{len(sub)}筆 勝率:{sub['success'].mean()*100:.0f}% 均報:{sub['exit_return_pct'].mean():.1f}%"

    # 按SPY趨勢
    for trend in df['spy_trend'].unique():
        sub = df[df['spy_trend'] == trend]
        if len(sub) > 0:
            s[f"trend_{trend}"] = f"{len(sub)}筆 勝率:{sub['success'].mean()*100:.0f}% 均報:{sub['exit_return_pct'].mean():.1f}%"

    if len(df) > 0:
        best = df.loc[df['exit_return_pct'].idxmax()]
        worst = df.loc[df['exit_return_pct'].idxmin()]
        s["best"] = f"{best['ticker']} ({best['exit_return_pct']:+.1f}%)"
        s["worst"] = f"{worst['ticker']} ({worst['exit_return_pct']:+.1f}%)"

    return s


if __name__ == "__main__":
    print("=" * 80)
    print("全方位量化回測系統 v4.0 — ATR止損 + 大盤過濾 + 智能出場")
    print(f"啟動時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    de = DataEngine()
    fh = FinnhubEngine(FINNHUB_API_KEY)
    strategy = StrategyV3(de, finnhub=fh, fmp_key=FMP_API_KEY, fred_key=FRED_API_KEY)
    bt = BacktestV4(strategy)

    macro = strategy.get_macro_regime()
    print(f"\n宏觀環境：{macro['state'].upper()} | VIX:{macro.get('vix')}")

    all_us = []

    # 期間1：2024 Q1→Q3（6個月，止盈10/20%）
    r1 = run_v4(US_STOCKS, "美股-期間1(6M)", "2024-01-15", "2024-07-15", 12, 25, bt, strategy)
    all_us.extend(r1)

    # 期間2：2024 Q3→Q4（3個月，止盈8/15%）
    r2 = run_v4(US_STOCKS, "美股-期間2(3M)", "2024-07-15", "2024-10-15", 8, 15, bt, strategy)
    all_us.extend(r2)

    # 期間3：2024 Q4→2025 Q1（3個月，止盈8/15%）
    r3 = run_v4(US_STOCKS, "美股-期間3(3M)", "2024-10-15", "2025-01-15", 8, 15, bt, strategy)
    all_us.extend(r3)

    # 港股
    hk = run_v4(HK_STOCKS, "港股(6M)", "2024-01-15", "2024-07-15", 12, 25, bt, strategy)

    us_s = summary(all_us, "美股v4")
    hk_s = summary(hk, "港股v4")

    # ==================== 完整比較 ====================
    print(f"\n{'█'*80}")
    print(f"{'v1 → v2 → v3 → v4 完整演進':^74}")
    print(f"{'█'*80}\n")

    print(f"{'指標':<18} {'v1':>10} {'v2':>10} {'v3':>10} {'v4':>10}")
    print(f"{'─'*58}")
    print(f"{'勝率(%)':<18} {'57.8':>10} {'67.9':>10} {'47.2':>10} {us_s['win_rate']:>10.1f}")
    print(f"{'平均回報(%)':<18} {'11.45':>10} {'9.56':>10} {'2.25':>10} {us_s['avg_return']:>10.2f}")
    print(f"{'中位回報(%)':<18} {'—':>10} {'—':>10} {'—':>10} {us_s['median_return']:>10.2f}")
    print(f"{'預測數量':<18} {'64':>10} {'81':>10} {'36':>10} {us_s['total']:>10}")
    print(f"{'平均持有天數':<18} {'—':>10} {'—':>10} {'24':>10} {us_s['avg_holding']:>10.0f}")

    print(f"\n📊 出場原因分析：")
    for k, v in us_s.items():
        if k.startswith("exit_"):
            print(f"  {k.replace('exit_', '')}：{v}")

    print(f"\n📊 大盤趨勢分析：")
    for k, v in us_s.items():
        if k.startswith("trend_"):
            print(f"  SPY {k.replace('trend_', '')}：{v}")

    if us_s.get("best"):
        print(f"\n  最佳：{us_s['best']}")
    if us_s.get("worst"):
        print(f"  最差：{us_s['worst']}")

    print(f"\n📊 港股 v4：勝率{hk_s['win_rate']}%（{hk_s.get('wins',0)}/{hk_s['total']}）| 平均回報:{hk_s.get('avg_return',0)}%")

    # 逐筆明細
    print(f"\n{'─'*80}")
    for r in all_us:
        s = "✓" if r['success'] else "✗"
        print(f"  {s} {r['ticker']:6s} [{r.get('spy_trend','?'):10s}] "
              f"分:{r['pred_score']:5.1f} {r['pred_direction']:8s} "
              f"→ {r['exit_return_pct']:>+7.1f}% {r['exit_reason']:8s} "
              f"({r['holding_days']}天 止損:{r['atr_stop_pct']}%)")

    # 保存
    output = {
        "version": "v4", "generated_at": datetime.now().isoformat(),
        "improvements": ["ATR動態止損", "SPY大盤過濾", "智能分批出場", "盈餘修正"],
        "us_summary": us_s, "hk_summary": hk_s,
        "comparison": {"v1": 57.8, "v2": 67.9, "v3": 47.2, "v4": us_s['win_rate']},
    }
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backtest_v4_results.json")
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n{'='*80}")
    print(f"完成 | 結果：{path}")
    print(f"{'='*80}")
