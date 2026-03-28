"""
回測引擎 — 執行歷史回測、驗證預測、追蹤績效
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from data_engine import DataEngine
from strategies import StrategyEngine


class Backtester:
    """回測引擎"""

    def __init__(self):
        self.de = DataEngine()
        self.se = StrategyEngine(self.de)
        self.trade_log = []
        self.prediction_log = []

    # ==================== 回測單支股票的預測準確度 ====================
    def backtest_prediction(self, ticker, prediction_date, direction,
                            expected_return_pct, timeframe_days, actual_end_date=None):
        """
        回測單筆預測是否準確

        參數：
            ticker: 股票代碼
            prediction_date: 預測日期
            direction: 預測方向（"漲" / "跌"）
            expected_return_pct: 預期回報百分比（如 15 表示 15%）
            timeframe_days: 預期實現天數
            actual_end_date: 實際觀察截止日（預設=prediction_date + timeframe_days）
        """
        if actual_end_date is None:
            end_dt = datetime.strptime(prediction_date, "%Y-%m-%d") + timedelta(days=timeframe_days + 30)
            actual_end_date = min(end_dt, datetime.now()).strftime("%Y-%m-%d")

        # 取得從預測日到觀察截止日的價格
        start_dt = datetime.strptime(prediction_date, "%Y-%m-%d") - timedelta(days=5)
        prices = self.de.get_prices(ticker, start_dt.strftime("%Y-%m-%d"), actual_end_date)

        if prices.empty:
            return {"error": f"無法獲取 {ticker} 的價格數據"}

        close = prices['Close']

        # 找到最接近預測日的價格
        pred_dt = pd.Timestamp(prediction_date)
        available_dates = close.index[close.index >= pred_dt]
        if len(available_dates) == 0:
            return {"error": "預測日期之後無數據"}

        entry_date = available_dates[0]
        entry_price = float(close.loc[entry_date])

        # 計算之後每天的回報
        future_prices = close.loc[entry_date:]
        returns = ((future_prices / entry_price) - 1) * 100

        # 判斷是否達到目標
        target_pct = expected_return_pct if direction == "漲" else -expected_return_pct
        target_price = entry_price * (1 + target_pct / 100)

        # 在時間框架內是否達標
        timeframe_end = entry_date + pd.Timedelta(days=timeframe_days)
        in_window = returns.loc[:timeframe_end]

        if direction == "漲":
            hit_target = (in_window >= expected_return_pct).any()
            max_favorable = float(in_window.max())
            max_adverse = float(in_window.min())
        else:
            hit_target = (in_window <= -expected_return_pct).any()
            max_favorable = float(-in_window.min())
            max_adverse = float(-in_window.max())

        if hit_target:
            # 找到首次達標日期
            if direction == "漲":
                hit_dates = in_window[in_window >= expected_return_pct].index
            else:
                hit_dates = in_window[in_window <= -expected_return_pct].index
            first_hit = hit_dates[0]
            days_to_hit = (first_hit - entry_date).days
        else:
            first_hit = None
            days_to_hit = None

        # 最終結果（觀察期結束時的回報）
        final_return = float(returns.iloc[-1])
        final_date = returns.index[-1]

        # 期間最大回撤
        cum_returns = future_prices / entry_price
        running_max = cum_returns.cummax()
        drawdowns = (cum_returns / running_max - 1) * 100
        max_drawdown = float(drawdowns.min())

        result = {
            "ticker": ticker,
            "prediction_date": prediction_date,
            "direction": direction,
            "expected_return_pct": expected_return_pct,
            "timeframe_days": timeframe_days,
            "entry_price": round(entry_price, 2),
            "target_price": round(target_price, 2),
            "hit_target": bool(hit_target),
            "days_to_hit": days_to_hit,
            "first_hit_date": str(first_hit.date()) if first_hit is not None else None,
            "max_favorable_pct": round(max_favorable, 2),
            "max_adverse_pct": round(max_adverse, 2),
            "max_drawdown_pct": round(max_drawdown, 2),
            "final_return_pct": round(final_return, 2),
            "final_date": str(final_date.date()),
            "final_price": round(float(future_prices.iloc[-1]), 2),
            "success": bool(hit_target),
        }

        # 分析原因
        result["analysis"] = self._analyze_result(ticker, prediction_date,
                                                   actual_end_date, result)

        self.prediction_log.append(result)
        return result

    def _analyze_result(self, ticker, start, end, result):
        """分析預測成功/失敗的原因"""
        analysis = {"reasons": []}

        try:
            # 取得技術信號
            tech = self.se.technical_signals(ticker, start, end)
            if not tech.empty:
                avg_signal = float(tech['total_signal'].mean())
                if result["direction"] == "漲" and avg_signal > 20:
                    analysis["reasons"].append("技術面持續看多，支撐了上漲")
                elif result["direction"] == "漲" and avg_signal < -20:
                    analysis["reasons"].append("技術面轉空，阻礙了上漲")

                # 均線狀態
                if 'SMA_200' in tech.columns:
                    prices = self.de.get_prices(ticker, start, end)
                    if not prices.empty:
                        close = prices['Close']
                        sma200 = close.rolling(200).mean()
                        if len(sma200.dropna()) > 0:
                            if close.iloc[-1] > sma200.iloc[-1]:
                                analysis["reasons"].append("價格在200日均線之上，長期趨勢向好")
                            else:
                                analysis["reasons"].append("價格在200日均線之下，長期趨勢偏弱")

            # 基本面變化
            info = self.de.get_stock_info(ticker)
            if info:
                eg = info.get("earningsGrowth")
                if eg:
                    if eg > 0.1:
                        analysis["reasons"].append(f"盈餘成長強勁({round(eg*100,1)}%)")
                    elif eg < -0.1:
                        analysis["reasons"].append(f"盈餘下滑({round(eg*100,1)}%)")

                rg = info.get("revenueGrowth")
                if rg:
                    if rg > 0.1:
                        analysis["reasons"].append(f"營收成長良好({round(rg*100,1)}%)")
                    elif rg < -0.05:
                        analysis["reasons"].append(f"營收衰退({round(rg*100,1)}%)")

            # 波動率分析
            if result["max_drawdown_pct"] < -15:
                analysis["reasons"].append(f"期間最大回撤達{result['max_drawdown_pct']}%，波動劇烈")

            if result["max_favorable_pct"] > result["expected_return_pct"] and not result["hit_target"]:
                analysis["reasons"].append("曾短暫達標但未持續，可能需要更好的出場時機")

        except Exception as e:
            analysis["reasons"].append(f"分析過程異常：{str(e)}")

        # 成功/失敗總結
        if result["success"]:
            analysis["verdict"] = "預測正確"
            analysis["lesson"] = "策略有效，信號可靠"
        else:
            if result["max_favorable_pct"] > result["expected_return_pct"] * 0.5:
                analysis["verdict"] = "方向正確但幅度/時間不足"
                analysis["lesson"] = "調整目標幅度或延長時間框架"
            elif result["final_return_pct"] * (1 if result["direction"] == "漲" else -1) > 0:
                analysis["verdict"] = "方向正確但未達目標"
                analysis["lesson"] = "降低目標回報或提高信號閾值"
            else:
                analysis["verdict"] = "預測錯誤"
                analysis["lesson"] = "需檢討入場信號強度與市場環境判斷"

        return analysis

    # ==================== 批量回測多支股票 ====================
    def batch_backtest(self, tickers, start_date, end_date,
                       expected_return=10, timeframe_days=90):
        """
        批量回測：對多支股票進行綜合分析 → 預測 → 驗證
        """
        results = []
        print(f"\n{'='*80}")
        print(f"批量回測開始：{len(tickers)} 支股票")
        print(f"回測期間：{start_date} → {end_date}")
        print(f"預期回報：{expected_return}% | 時間框架：{timeframe_days} 天")
        print(f"{'='*80}\n")

        for i, ticker in enumerate(tickers):
            print(f"[{i+1}/{len(tickers)}] 分析 {ticker}...", end=" ")
            try:
                # 綜合分析
                analysis = self.se.comprehensive_analysis(ticker, start_date, end_date)
                pred = analysis.get("prediction", {})

                if pred.get("direction") in ["強烈看漲", "看漲"]:
                    # 執行回測驗證
                    bt = self.backtest_prediction(
                        ticker=ticker,
                        prediction_date=start_date,
                        direction="漲",
                        expected_return_pct=expected_return,
                        timeframe_days=timeframe_days,
                        actual_end_date=end_date,
                    )
                    bt["pred_direction"] = pred["direction"]
                    bt["pred_score"] = pred.get("score", 0)
                    bt["pred_confidence"] = pred.get("confidence", "")
                    bt["bull_reasons"] = pred.get("bull_reasons", [])
                    bt["bear_reasons"] = pred.get("bear_reasons", [])
                    results.append(bt)
                    status = "✓ 達標" if bt.get("success") else "✗ 未達標"
                    print(f"{pred['direction']} → {status} (實際:{bt.get('final_return_pct',0)}%)")
                elif pred.get("direction") in ["強烈看跌", "看跌"]:
                    bt = self.backtest_prediction(
                        ticker=ticker,
                        prediction_date=start_date,
                        direction="跌",
                        expected_return_pct=expected_return,
                        timeframe_days=timeframe_days,
                        actual_end_date=end_date,
                    )
                    bt["pred_direction"] = pred["direction"]
                    bt["pred_score"] = pred.get("score", 0)
                    bt["pred_confidence"] = pred.get("confidence", "")
                    bt["bull_reasons"] = pred.get("bull_reasons", [])
                    bt["bear_reasons"] = pred.get("bear_reasons", [])
                    results.append(bt)
                    status = "✓ 達標" if bt.get("success") else "✗ 未達標"
                    print(f"{pred['direction']} → {status} (實際:{bt.get('final_return_pct',0)}%)")
                else:
                    print(f"觀望（評分:{pred.get('score',0)}）")

            except Exception as e:
                print(f"錯誤: {str(e)[:50]}")
                continue

        return results

    # ==================== 績效統計 ====================
    def performance_summary(self, results):
        """計算回測績效統計"""
        if not results:
            return "無回測結果"

        df = pd.DataFrame(results)

        total = len(df)
        hits = df['success'].sum() if 'success' in df.columns else 0
        win_rate = hits / total * 100 if total > 0 else 0

        # 按方向分類
        summary = {
            "total_predictions": total,
            "successful": int(hits),
            "failed": total - int(hits),
            "win_rate": round(win_rate, 1),
            "avg_final_return": round(df['final_return_pct'].mean(), 2) if 'final_return_pct' in df.columns else 0,
            "avg_max_favorable": round(df['max_favorable_pct'].mean(), 2) if 'max_favorable_pct' in df.columns else 0,
            "avg_max_drawdown": round(df['max_drawdown_pct'].mean(), 2) if 'max_drawdown_pct' in df.columns else 0,
            "best_trade": None,
            "worst_trade": None,
        }

        if 'final_return_pct' in df.columns and len(df) > 0:
            best_idx = df['final_return_pct'].idxmax()
            worst_idx = df['final_return_pct'].idxmin()
            summary["best_trade"] = f"{df.loc[best_idx, 'ticker']} (+{df.loc[best_idx, 'final_return_pct']}%)"
            summary["worst_trade"] = f"{df.loc[worst_idx, 'ticker']} ({df.loc[worst_idx, 'final_return_pct']}%)"

        # 高信心 vs 低信心的勝率
        if 'pred_score' in df.columns:
            high_conf = df[df['pred_score'] >= 5]
            low_conf = df[df['pred_score'] < 5]
            if len(high_conf) > 0:
                summary["high_confidence_win_rate"] = round(high_conf['success'].mean() * 100, 1)
            if len(low_conf) > 0:
                summary["low_confidence_win_rate"] = round(low_conf['success'].mean() * 100, 1)

        return summary

    def print_report(self, results, summary):
        """打印完整報告"""
        print(f"\n{'='*80}")
        print(f"{'回測績效報告':^76}")
        print(f"{'='*80}")

        print(f"\n📊 總覽")
        print(f"  總預測數：{summary['total_predictions']}")
        print(f"  成功：{summary['successful']} | 失敗：{summary['failed']}")
        print(f"  勝率：{summary['win_rate']}%")
        print(f"  平均最終回報：{summary['avg_final_return']}%")
        print(f"  平均最大有利：{summary['avg_max_favorable']}%")
        print(f"  平均最大回撤：{summary['avg_max_drawdown']}%")

        if summary.get("best_trade"):
            print(f"  最佳交易：{summary['best_trade']}")
        if summary.get("worst_trade"):
            print(f"  最差交易：{summary['worst_trade']}")

        if summary.get("high_confidence_win_rate") is not None:
            print(f"\n  高信心預測勝率：{summary['high_confidence_win_rate']}%")
        if summary.get("low_confidence_win_rate") is not None:
            print(f"  低信心預測勝率：{summary['low_confidence_win_rate']}%")

        print(f"\n{'─'*80}")
        print(f"{'逐筆交易明細':^76}")
        print(f"{'─'*80}")

        for r in results:
            status = "✓" if r.get('success') else "✗"
            ticker = r.get('ticker', '?')
            direction = r.get('pred_direction', '?')
            final = r.get('final_return_pct', 0)
            max_fav = r.get('max_favorable_pct', 0)
            max_dd = r.get('max_drawdown_pct', 0)

            print(f"\n  {status} {ticker:8s} | 預測:{direction:6s} | "
                  f"最終:{final:>7.1f}% | 最大有利:{max_fav:>7.1f}% | 最大回撤:{max_dd:>7.1f}%")

            # 原因分析
            analysis = r.get('analysis', {})
            if analysis.get('reasons'):
                for reason in analysis['reasons'][:3]:
                    print(f"    → {reason}")
            if analysis.get('verdict'):
                print(f"    📋 結論：{analysis['verdict']}")
            if analysis.get('lesson'):
                print(f"    💡 教訓：{analysis['lesson']}")

        print(f"\n{'='*80}\n")
