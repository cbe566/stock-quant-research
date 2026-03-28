"""
策略引擎 v3 — 目標勝率75%+
在 v2 基礎上新增：
P0-1: 盈餘修正因子（Earnings Revision）
P0-2: Triple Barrier 出場策略（止盈+止損+時間限制）
P0-3: 提高交易門檻 + 信心加權倉位
"""

import pandas as pd
import numpy as np
import requests
from datetime import datetime, timedelta
from strategies_v2 import StrategyV2, SECTOR_MAP
from data_engine import DataEngine
from config import FMP_API_KEY, FINNHUB_API_KEY


class StrategyV3(StrategyV2):
    """v3 策略引擎：繼承 v2 全部功能，新增三大P0改善"""

    def __init__(self, data_engine, finnhub=None, fmp_key=None, fred_key=None):
        super().__init__(data_engine, finnhub, fmp_key, fred_key)
        self._revision_cache = {}

    # ==================== P0-1: 盈餘修正因子 ====================
    def earnings_revision_score(self, ticker):
        """
        盈餘修正因子：分析師是否上調/下調盈餘預期
        學術IC: 0.04-0.07，為最有效的單因子之一
        """
        reasons = []
        score = 50  # 中性起點

        if not self.fmp_key:
            return score, reasons

        try:
            base = "https://financialmodelingprep.com/stable"

            # 獲取分析師預估（季度）
            r = requests.get(
                f"{base}/analyst-estimates?symbol={ticker}&limit=8&apikey={self.fmp_key}",
                timeout=10
            )
            if r.status_code != 200 or not r.text.strip():
                return score, reasons

            data = r.json()
            if not data or not isinstance(data, list) or len(data) < 2:
                return score, reasons

            # 比較最近兩期的預估值變化
            latest = data[0]
            previous = data[1]

            # EPS 預估變化
            eps_latest = latest.get("estimatedEpsAvg", 0)
            eps_prev = previous.get("estimatedEpsAvg", 0)

            # 營收預估變化
            rev_latest = latest.get("estimatedRevenueAvg", 0)
            rev_prev = previous.get("estimatedRevenueAvg", 0)

            # 計算修正幅度
            if eps_prev and eps_prev != 0:
                eps_revision = (eps_latest - eps_prev) / abs(eps_prev) * 100
            else:
                eps_revision = 0

            if rev_prev and rev_prev != 0:
                rev_revision = (rev_latest - rev_prev) / abs(rev_prev) * 100
            else:
                rev_revision = 0

            # 評分邏輯
            # EPS 修正（權重更大）
            if eps_revision > 10:
                score += 25
                reasons.append(f"EPS預估大幅上調({eps_revision:+.1f}%)")
            elif eps_revision > 3:
                score += 15
                reasons.append(f"EPS預估上調({eps_revision:+.1f}%)")
            elif eps_revision < -10:
                score -= 25
                reasons.append(f"EPS預估大幅下調({eps_revision:+.1f}%)")
            elif eps_revision < -3:
                score -= 15
                reasons.append(f"EPS預估下調({eps_revision:+.1f}%)")

            # 營收修正
            if rev_revision > 5:
                score += 10
                reasons.append(f"營收預估上調({rev_revision:+.1f}%)")
            elif rev_revision < -5:
                score -= 10
                reasons.append(f"營收預估下調({rev_revision:+.1f}%)")

            # 預估數量（分析師覆蓋度）
            num_analysts = latest.get("numberAnalystEstimatedEps", 0)
            if num_analysts and num_analysts > 20:
                # 高覆蓋度的修正信號更可靠
                if abs(eps_revision) > 5:
                    bonus = 5 if eps_revision > 0 else -5
                    score += bonus

        except Exception:
            pass

        score = max(min(score, 100), 0)
        return score, reasons

    # ==================== P0-2: Triple Barrier 出場策略 ====================
    @staticmethod
    def triple_barrier_exit(prices_series, entry_price, take_profit_pct=15,
                            stop_loss_pct=8, max_holding_days=90):
        """
        Triple Barrier 出場策略：
        - 上界：獲利 take_profit_pct% → 止盈
        - 下界：虧損 stop_loss_pct% → 止損
        - 時間界：持有 max_holding_days 天 → 強制平倉

        回傳：(exit_price, exit_date, exit_reason, return_pct, holding_days)
        """
        upper = entry_price * (1 + take_profit_pct / 100)
        lower = entry_price * (1 - stop_loss_pct / 100)

        for i, (date, price) in enumerate(prices_series.items()):
            price_val = float(price)
            days_held = i

            # 止盈觸發
            if price_val >= upper:
                ret = (price_val / entry_price - 1) * 100
                return {
                    "exit_price": round(price_val, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": "止盈",
                    "return_pct": round(ret, 2),
                    "holding_days": days_held,
                }

            # 止損觸發
            if price_val <= lower:
                ret = (price_val / entry_price - 1) * 100
                return {
                    "exit_price": round(price_val, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": "止損",
                    "return_pct": round(ret, 2),
                    "holding_days": days_held,
                }

            # 時間限制
            if days_held >= max_holding_days:
                ret = (price_val / entry_price - 1) * 100
                return {
                    "exit_price": round(price_val, 2),
                    "exit_date": str(date.date()) if hasattr(date, 'date') else str(date),
                    "exit_reason": "時間到期",
                    "return_pct": round(ret, 2),
                    "holding_days": days_held,
                }

        # 數據不夠
        if len(prices_series) > 0:
            last_price = float(prices_series.iloc[-1])
            ret = (last_price / entry_price - 1) * 100
            return {
                "exit_price": round(last_price, 2),
                "exit_date": str(prices_series.index[-1].date()) if hasattr(prices_series.index[-1], 'date') else "",
                "exit_reason": "數據截止",
                "return_pct": round(ret, 2),
                "holding_days": len(prices_series),
            }

        return {"exit_reason": "無數據", "return_pct": 0, "holding_days": 0}

    # ==================== P0-3: 信心等級與倉位管理 ====================
    @staticmethod
    def confidence_level(total_score, tech_score, revision_score):
        """
        信心分級：決定是否出手 + 倉位大小
        A級（強信心）：滿倉出手
        B級（中信心）：半倉出手
        C級（弱信心）：不出手
        """
        # 基礎信心 = 總分偏離中性的程度
        deviation = abs(total_score - 50)

        # 技術面確認加分
        if (total_score > 60 and tech_score > 20) or (total_score < 40 and tech_score < -20):
            deviation += 10  # 方向一致 → 加信心

        # 盈餘修正確認加分
        if (total_score > 60 and revision_score > 60) or (total_score < 40 and revision_score < 40):
            deviation += 8  # 盈餘修正方向一致 → 加信心

        # 技術面/基本面衝突減分
        if (total_score > 60 and tech_score < -20) or (total_score < 40 and tech_score > 20):
            deviation -= 15  # 方向衝突 → 減信心

        if deviation >= 18:
            return "A", 1.0    # 全倉（高信心）
        elif deviation >= 8:
            return "B", 0.5    # 半倉（中信心）
        else:
            return "C", 0.0    # 不出手（信心不足）

    # ==================== v3 綜合分析（覆寫 v2）====================
    def comprehensive_analysis_v3(self, ticker, start, end):
        """
        v3 綜合分析：在 v2 基礎上加入盈餘修正、信心過濾
        """
        # 先跑 v2 分析
        result = self.comprehensive_analysis_v2(ticker, start, end)

        # P0-1: 加入盈餘修正因子
        rev_score, rev_reasons = self.earnings_revision_score(ticker)
        result["revision_score"] = round(rev_score, 1)

        # 將盈餘修正納入總分（權重 10%，從其他因子各減少一點）
        old_total = result["total_score"]
        # 重新計算：原始分數占90%，盈餘修正占10%
        new_total = old_total * 0.88 + rev_score * 0.12
        result["total_score"] = round(new_total, 1)

        # P0-3: 信心分級
        tech_raw = result.get("tech_score", 50) * 2 - 100  # 轉回 -100~100
        conf_level, position_size = self.confidence_level(
            new_total, tech_raw, rev_score
        )
        result["confidence_level"] = conf_level
        result["position_size"] = position_size

        # 更新預測（用新的總分和信心）
        all_bull = result.get("prediction", {}).get("bull_reasons", [])
        all_bear = result.get("prediction", {}).get("bear_reasons", [])
        all_bull.extend([r for r in rev_reasons if "上調" in r])
        all_bear.extend([r for r in rev_reasons if "下調" in r])

        result["prediction"] = self._generate_prediction_v3(
            new_total, tech_raw, rev_score, conf_level,
            all_bull, all_bear, ticker,
            self.get_macro_regime()
        )

        return result

    def _generate_prediction_v3(self, total_score, tech_score, rev_score,
                                 conf_level, reasons_bull, reasons_bear,
                                 ticker, macro):
        """v3 預測生成：加入信心過濾"""

        # P0-3: C級信心直接觀望
        if conf_level == "C":
            return {
                "direction": "觀望（信心不足）",
                "total_score": round(total_score, 1),
                "tech_score": tech_score,
                "revision_score": rev_score,
                "confidence": conf_level,
                "position_size": 0,
                "expected_return": "—",
                "timeframe": "—",
                "macro_regime": macro.get("state"),
                "bull_reasons": reasons_bull[:3],
                "bear_reasons": reasons_bear[:3],
            }

        # 根據環境調整閾值
        state = macro.get("state", "neutral")
        if state == "risk_off":
            buy_threshold = 65
            strong_buy = 78
        elif state == "risk_on":
            buy_threshold = 56
            strong_buy = 70
        else:
            buy_threshold = 60
            strong_buy = 74

        # 盈餘修正可以調整閾值
        if rev_score > 70:
            buy_threshold -= 3  # 盈餘上調降低買入門檻
        elif rev_score < 30:
            buy_threshold += 3  # 盈餘下調提高門檻

        if total_score >= strong_buy:
            direction = "強烈看漲"
            expected = "15-25%"
            tf = "3-6個月"
        elif total_score >= buy_threshold:
            direction = "看漲"
            expected = "8-15%"
            tf = "3-6個月"
        elif total_score >= 48:
            direction = "輕度看漲"
            expected = "3-8%"
            tf = "1-3個月"
        elif total_score >= 38:
            direction = "中性/觀望"
            expected = "-3%~3%"
            tf = "—"
        elif total_score >= 28:
            direction = "看跌"
            expected = "-5%~-15%"
            tf = "1-3個月"
        else:
            direction = "強烈看跌"
            expected = "-15%以下"
            tf = "1-3個月"

        # B級信心降級處理
        if conf_level == "B" and direction in ["強烈看漲", "強烈看跌"]:
            direction = direction.replace("強烈", "")  # 降級

        return {
            "direction": direction,
            "total_score": round(total_score, 1),
            "tech_score": tech_score,
            "revision_score": rev_score,
            "confidence": conf_level,
            "position_size": 1.0 if conf_level == "A" else 0.5,
            "expected_return": expected,
            "timeframe": tf,
            "macro_regime": macro.get("state"),
            "bull_reasons": reasons_bull[:5],
            "bear_reasons": reasons_bear[:5],
        }
