#!/usr/bin/env python3
"""
最終回測系統 — 無前視偏差版
使用歷史時點數據：SimFin歷史財報 + Tiingo歷史股價 + FRED歷史宏觀 + Marketaux歷史新聞
"""

import sys, os, json
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import US_STOCKS
from historical_data_engine import HistoricalDataEngine
import pandas as pd
import numpy as np


class FinalBacktest:
    """無前視偏差的最終回測"""

    def __init__(self):
        self.hde = HistoricalDataEngine()

    def score_stock(self, ticker, as_of_date):
        """
        用歷史數據評分（只用as_of_date當天能看到的數據）
        """
        scores = {}
        reasons_bull = []
        reasons_bear = []

        # 1. 歷史財報
        fin = self.hde.get_historical_financials(ticker, as_of_date)
        val = self.hde.calc_historical_valuation(ticker, as_of_date, fin)

        # 質量因子
        q = 0
        if fin:
            roe = fin.get("roe", 0)
            if roe > 0.20: q += 30; reasons_bull.append(f"ROE高({roe:.0%})")
            elif roe > 0.10: q += 15
            elif roe < 0: q -= 20; reasons_bear.append(f"ROE為負({roe:.0%})")

            gm = fin.get("gross_margin", 0)
            if gm > 0.50: q += 20
            elif gm > 0.30: q += 10

            de = fin.get("debt_equity", 0)
            if 0 < de < 1.0: q += 15
            elif de > 2.0: q -= 15; reasons_bear.append(f"負債過高(D/E:{de:.1f})")

            fcf = fin.get("fcf", 0)
            if fcf > 0: q += 15;
            else: q -= 10

            nm = fin.get("net_margin", 0)
            if nm > 0.20: q += 20
            elif nm > 0.10: q += 10
        scores["quality"] = max(min(q, 100), 0)

        # 價值因子
        v = 50  # 中性起點
        if val:
            pe = val.get("pe_ttm")
            if pe:
                if 0 < pe < 15: v += 30; reasons_bull.append(f"PE低估({pe:.0f})")
                elif 0 < pe < 25: v += 10
                elif pe > 50: v -= 25; reasons_bear.append(f"PE過高({pe:.0f})")
                elif pe > 35: v -= 10

            pb = val.get("pb")
            if pb:
                if 0 < pb < 2: v += 15
                elif 0 < pb < 4: v += 5
                elif pb > 10: v -= 15
        scores["value"] = max(min(v, 100), 0)

        # 2. 技術面
        tech = self.hde.calc_historical_technicals(ticker, as_of_date)
        t = 50
        if tech:
            # 趨勢
            if tech.get("above_sma200") and tech.get("sma50_above_sma200"):
                t += 20; reasons_bull.append("強勢趨勢(>50MA>200MA)")
            elif tech.get("above_sma200"):
                t += 10
            elif not tech.get("above_sma200") and not tech.get("sma50_above_sma200"):
                t -= 20; reasons_bear.append("弱勢趨勢(<50MA<200MA)")

            # RSI
            rsi = tech.get("rsi", 50)
            if rsi < 30: t += 15; reasons_bull.append(f"RSI超賣({rsi:.0f})")
            elif rsi > 70: t -= 15; reasons_bear.append(f"RSI超買({rsi:.0f})")

            # MACD
            if tech.get("macd_cross_up"): t += 10; reasons_bull.append("MACD金叉")
            elif tech.get("macd_cross_down"): t -= 10; reasons_bear.append("MACD死叉")
            elif tech.get("macd_hist", 0) > 0: t += 5
            elif tech.get("macd_hist", 0) < 0: t -= 5

            # 動量
            mom6 = tech.get("mom_6m", 0)
            if mom6 > 20: t += 10; reasons_bull.append(f"6月動量+{mom6:.0f}%")
            elif mom6 < -15: t -= 10; reasons_bear.append(f"6月動量{mom6:.0f}%")

            # 量比
            vr = tech.get("vol_ratio", 1)
            if vr > 2.0: t += 5
        scores["technical"] = max(min(t, 100), 0)

        # 3. 宏觀環境
        macro = self.hde.get_macro_at_date(as_of_date)
        m = 50
        regime = macro.get("regime", "neutral")
        if regime == "risk_on": m = 70
        elif regime == "risk_off": m = 30
        scores["macro"] = m

        # 4. 新聞情緒（Marketaux限額有限，可選）
        # news = self.hde.get_news_sentiment(ticker, as_of_date)
        # scores["sentiment"] = ...
        # 暫時不用，避免消耗API額度

        # ==================== 加權總分 ====================
        weights = {"quality": 0.30, "value": 0.25, "technical": 0.30, "macro": 0.15}
        total = sum(scores[k] * weights[k] for k in weights)

        # 宏觀環境修正
        if regime == "risk_off" and total > 60:
            total *= 0.90  # 避險環境降低信心
        elif regime == "risk_on" and total > 50:
            total *= 1.05

        total = max(min(total, 100), 0)

        # 預測方向
        if total >= 65:
            direction = "看漲"
        elif total >= 55:
            direction = "輕度看漲"
        elif total >= 45:
            direction = "觀望"
        elif total >= 35:
            direction = "輕度看跌"
        else:
            direction = "看跌"

        return {
            "ticker": ticker,
            "as_of_date": as_of_date,
            "total_score": round(total, 1),
            "quality": scores["quality"],
            "value": scores["value"],
            "technical": scores["technical"],
            "macro": scores["macro"],
            "regime": regime,
            "direction": direction,
            "bull_reasons": reasons_bull[:4],
            "bear_reasons": reasons_bear[:4],
            "price": tech.get("price") if tech else None,
        }

    def backtest_trade(self, ticker, entry_date, exit_date, direction, take_profit=15, hard_stop=20):
        """回測一筆交易"""
        prices = self.hde.get_historical_prices(ticker, entry_date, exit_date)
        if prices.empty:
            return None

        close = prices['Close']
        entry_price = float(close.iloc[0])

        # 止盈/硬止損出場
        for i, (dt, px) in enumerate(close.items()):
            pv = float(px)
            if direction in ["看漲", "輕度看漲"]:
                ret = (pv / entry_price - 1) * 100
                if ret >= take_profit:
                    return self._make_result(ticker, entry_price, pv, ret, i, "止盈", close)
                if ret <= -hard_stop:
                    return self._make_result(ticker, entry_price, pv, ret, i, "硬止損", close)
            elif direction in ["看跌", "輕度看跌"]:
                ret = (entry_price / pv - 1) * 100
                if ret >= take_profit:
                    return self._make_result(ticker, entry_price, pv, ret, i, "止盈", close, short=True)
                if ret <= -hard_stop:
                    return self._make_result(ticker, entry_price, pv, ret, i, "硬止損", close, short=True)

        # 持有到期
        last = float(close.iloc[-1])
        if direction in ["看漲", "輕度看漲"]:
            ret = (last / entry_price - 1) * 100
        else:
            ret = (entry_price / last - 1) * 100
        return self._make_result(ticker, entry_price, last, ret, len(close)-1, "持有到期", close,
                                 short="看跌" in direction)

    def _make_result(self, ticker, entry, exit_px, ret, days, reason, close, short=False):
        returns = ((close / entry) - 1) * 100
        return {
            "ticker": ticker,
            "entry_price": round(entry, 2),
            "exit_price": round(exit_px, 2),
            "return_pct": round(ret, 2),
            "holding_days": days,
            "exit_reason": reason,
            "success": ret > 0,
            "max_favorable": round(float(returns.max() if not short else -returns.min()), 2),
            "max_drawdown": round(float(returns.min() if not short else -returns.max()), 2),
        }


def run_period(bt, tickers, label, entry_date, exit_date, tp, stop):
    """跑一個期間"""
    results = []
    print(f"\n{'='*80}")
    print(f"[無偏差回測] {label} | {entry_date}→{exit_date} | 止盈{tp}%/止損{stop}%")

    # 取宏觀環境
    macro = bt.hde.get_macro_at_date(entry_date)
    print(f"宏觀環境：{macro.get('regime','?')} | VIX:{macro.get('vix','?')} | SPY趨勢:{macro.get('spy_trend','?')}")
    print(f"{'='*80}\n")

    for i, t in enumerate(tickers):
        print(f"[{i+1}/{len(tickers)}] {t}...", end=" ")
        try:
            score = bt.score_stock(t, entry_date)
            direction = score["direction"]
            total = score["total_score"]

            if direction == "觀望":
                print(f"觀望(分:{total:.0f})")
                continue

            # 大盤下跌時不做多
            if "看漲" in direction and macro.get("spy_trend") == "down":
                print(f"{direction}(分:{total:.0f}) — 大盤弱勢過濾")
                continue

            trade = bt.backtest_trade(t, entry_date, exit_date, direction, tp, stop)
            if trade:
                trade["pred_direction"] = direction
                trade["pred_score"] = total
                trade["quality"] = score["quality"]
                trade["value"] = score["value"]
                trade["technical"] = score["technical"]
                trade["macro_score"] = score["macro"]
                trade["regime"] = score["regime"]
                trade["bull_reasons"] = score["bull_reasons"]
                trade["bear_reasons"] = score["bear_reasons"]
                results.append(trade)

                s = "✓" if trade["success"] else "✗"
                print(f"{s} {direction}(分:{total:.0f}) → {trade['return_pct']:+.1f}% "
                      f"({trade['exit_reason']}, {trade['holding_days']}天)")
            else:
                print(f"{direction}(分:{total:.0f}) — 無數據")

        except Exception as e:
            print(f"錯誤:{str(e)[:40]}")

    return results


if __name__ == "__main__":
    print("█" * 80)
    print("最終回測系統 — 無前視偏差版")
    print(f"使用歷史時點數據：SimFin + Tiingo + FRED")
    print(f"時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("█" * 80)

    bt = FinalBacktest()
    all_results = []

    # 期間1：2024 Q1→Q3（牛市）
    r1 = run_period(bt, US_STOCKS, "P1 牛市(6M)", "2024-01-15", "2024-07-15", 15, 20)
    all_results.extend(r1)

    # 期間2：2024 Q3→Q4（回調）
    r2 = run_period(bt, US_STOCKS, "P2 回調(3M)", "2024-07-15", "2024-10-15", 12, 18)
    all_results.extend(r2)

    # 期間3：2024 Q4→2025 Q1（震盪）
    r3 = run_period(bt, US_STOCKS, "P3 震盪(3M)", "2024-10-15", "2025-01-15", 12, 18)
    all_results.extend(r3)

    # ==================== 統計 ====================
    if all_results:
        df = pd.DataFrame(all_results)
        total = len(df)
        wins = int(df['success'].sum())
        wr = round(wins / total * 100, 1)

        wins_df = df[df['return_pct'] > 0]
        loss_df = df[df['return_pct'] <= 0]
        avg_win = wins_df['return_pct'].mean() if len(wins_df) > 0 else 0
        avg_loss = abs(loss_df['return_pct'].mean()) if len(loss_df) > 0 else 1
        pf = wins_df['return_pct'].sum() / abs(loss_df['return_pct'].sum()) if len(loss_df) > 0 and loss_df['return_pct'].sum() != 0 else 999

        print(f"\n{'█'*80}")
        print(f"{'最終結果（無前視偏差）':^74}")
        print(f"{'█'*80}\n")

        print(f"  勝率：{wr}%（{wins}/{total}）")
        print(f"  平均回報：{df['return_pct'].mean():.2f}%")
        print(f"  中位回報：{df['return_pct'].median():.2f}%")
        print(f"  平均贏利：+{avg_win:.2f}% | 平均虧損：-{avg_loss:.2f}%")
        print(f"  盈虧比：{avg_win/avg_loss:.2f}")
        print(f"  利潤因子：{pf:.2f}")
        print(f"  期望值：{wr/100*avg_win - (1-wr/100)*avg_loss:.2f}%/筆")
        print(f"  平均持有：{df['holding_days'].mean():.0f}天")

        # 出場原因
        print(f"\n  出場分析：")
        for reason in df['exit_reason'].unique():
            sub = df[df['exit_reason'] == reason]
            print(f"    {reason}：{len(sub)}筆 勝率:{sub['success'].mean()*100:.0f}% 均報:{sub['return_pct'].mean():.1f}%")

        # 按期間
        print(f"\n  按期間：")
        for i, (r, label) in enumerate([(r1, "P1牛市"), (r2, "P2回調"), (r3, "P3震盪")]):
            if r:
                sub = pd.DataFrame(r)
                print(f"    {label}：{len(sub)}筆 勝率:{sub['success'].mean()*100:.0f}% 均報:{sub['return_pct'].mean():.1f}%")

        # 逐筆
        print(f"\n{'─'*80}")
        for r in all_results:
            s = "✓" if r['success'] else "✗"
            print(f"  {s} {r['ticker']:6s} | {r.get('pred_direction','?'):6s} 分:{r.get('pred_score',0):5.1f} "
                  f"| {r['return_pct']:>+7.1f}% | {r['exit_reason']:6s} {r['holding_days']:>3}天 "
                  f"| Q:{r.get('quality',0):>3} V:{r.get('value',0):>3} T:{r.get('technical',0):>3}")
            for reason in r.get('bull_reasons', [])[:1]:
                print(f"       ↑ {reason}")
            for reason in r.get('bear_reasons', [])[:1]:
                print(f"       ↓ {reason}")

        if len(df) > 0:
            print(f"\n  最佳：{df.loc[df['return_pct'].idxmax(), 'ticker']} ({df['return_pct'].max():+.1f}%)")
            print(f"  最差：{df.loc[df['return_pct'].idxmin(), 'ticker']} ({df['return_pct'].min():+.1f}%)")

        # 保存
        output = {
            "version": "final_no_bias",
            "generated_at": datetime.now().isoformat(),
            "data_sources": "SimFin歷史財報 + Tiingo歷史股價 + FRED歷史宏觀",
            "win_rate": wr, "avg_return": round(df['return_pct'].mean(), 2),
            "profit_factor": round(pf, 2), "total_trades": total,
        }
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backtest_final_results.json")
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2, default=str)
        print(f"\n結果：{path}")

    print(f"\n{'='*80}")
    print(f"完成：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*80}")
