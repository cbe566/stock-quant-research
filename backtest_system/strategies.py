"""
策略引擎 — 實現所有策略的選股與信號生成
包含：多因子、技術面、動量、均值回歸、綜合評分
"""

import pandas as pd
import numpy as np
from data_engine import DataEngine


class StrategyEngine:
    """策略引擎：生成買賣信號與預測"""

    def __init__(self, data_engine: DataEngine):
        self.de = data_engine

    # ==================== 策略一：多因子選股 ====================
    def multifactor_score(self, ticker, params):
        """
        QVM 多因子評分（Quality-Value-Momentum）
        回傳：0-100 的綜合評分
        """
        info = self.de.get_stock_info(ticker)
        if not info:
            return None

        scores = {}
        weights = {"quality": 0.35, "value": 0.35, "momentum": 0.30}

        # --- 質量因子（Quality）---
        q_score = 0
        q_max = 5

        roe = info.get("returnOnEquity", None)
        if roe is not None:
            if roe > 0.20: q_score += 2
            elif roe > 0.10: q_score += 1

        gm = info.get("grossMargins", None)
        if gm is not None:
            if gm > 0.40: q_score += 1.5
            elif gm > 0.25: q_score += 0.75

        dte = info.get("debtToEquity", None)
        if dte is not None:
            if dte < 50: q_score += 1
            elif dte < 100: q_score += 0.5

        fcf = info.get("freeCashflow", None)
        if fcf is not None and fcf > 0:
            q_score += 0.5

        scores["quality"] = min(q_score / q_max, 1.0) * 100

        # --- 價值因子（Value）---
        v_score = 0
        v_max = 5

        pe = info.get("forwardPE", None) or info.get("trailingPE", None)
        if pe is not None:
            if 0 < pe < 15: v_score += 2
            elif 0 < pe < 25: v_score += 1

        pb = info.get("priceToBook", None)
        if pb is not None:
            if 0 < pb < 2: v_score += 1.5
            elif 0 < pb < 5: v_score += 0.75

        peg = info.get("pegRatio", None)
        if peg is not None:
            if 0 < peg < 1: v_score += 1
            elif 0 < peg < 2: v_score += 0.5

        dy = info.get("dividendYield", None)
        if dy is not None and dy > 0.02:
            v_score += 0.5

        scores["value"] = min(v_score / v_max, 1.0) * 100

        # --- 動量因子（Momentum）---
        m_score = 0
        m_max = 3

        try:
            from datetime import datetime, timedelta
            end = datetime.now().strftime("%Y-%m-%d")
            start = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
            prices = self.de.get_prices(ticker, start, end)

            if len(prices) > 126:
                close = prices['Close']
                # 6個月動量（排除最近1個月）
                mom_6m = (close.iloc[-22] / close.iloc[-148]) - 1 if len(close) > 148 else 0
                if mom_6m > 0.15: m_score += 1.5
                elif mom_6m > 0.05: m_score += 0.75

                # 是否在200日均線之上
                sma200 = close.rolling(200).mean()
                if len(sma200.dropna()) > 0 and close.iloc[-1] > sma200.iloc[-1]:
                    m_score += 1

                # RSI 不過熱
                rsi = self.de.calc_rsi(close, 14)
                if 40 < rsi.iloc[-1] < 70:
                    m_score += 0.5

        except Exception:
            pass

        scores["momentum"] = min(m_score / m_max, 1.0) * 100

        # --- 綜合評分 ---
        total = (scores.get("quality", 0) * weights["quality"] +
                 scores.get("value", 0) * weights["value"] +
                 scores.get("momentum", 0) * weights["momentum"])

        return {
            "ticker": ticker,
            "total_score": round(total, 1),
            "quality_score": round(scores.get("quality", 0), 1),
            "value_score": round(scores.get("value", 0), 1),
            "momentum_score": round(scores.get("momentum", 0), 1),
            "f_score": self.de.calc_piotroski_f_score(ticker).get("score", 0),
            "recommendation": self._get_recommendation(total),
        }

    @staticmethod
    def _get_recommendation(score):
        if score >= 75: return "強烈買入"
        elif score >= 60: return "買入"
        elif score >= 45: return "觀望"
        elif score >= 30: return "減持"
        else: return "賣出"

    # ==================== 策略二：技術面綜合信號 ====================
    def technical_signals(self, ticker, start, end, params=None):
        """
        生成技術面買賣信號
        回傳每天的信號強度（-100 到 +100）
        """
        prices = self.de.get_prices(ticker, start, end)
        if prices.empty:
            return pd.DataFrame()

        df = self.de.calc_all_technicals(prices)
        close = df['Close']

        signals = pd.DataFrame(index=df.index)
        signals['close'] = close

        # 1. 均線系統（權重30%）
        ma_signal = pd.Series(0.0, index=df.index)
        # 黃金交叉/死亡交叉
        ma_signal += np.where(df['SMA_20'] > df['SMA_50'], 15, -15)
        ma_signal += np.where(df['SMA_50'] > df['SMA_200'], 15, -15)
        # 價格在均線之上/下
        ma_signal = ma_signal.astype(float)
        ma_signal += np.where(close > df['SMA_200'], 10, -10)
        signals['ma_signal'] = ma_signal.clip(-30, 30)

        # 2. RSI（權重20%）
        rsi = df['RSI_14']
        rsi_signal = pd.Series(0.0, index=df.index)
        rsi_signal = np.where(rsi < 30, 20, np.where(rsi > 70, -20,
                              np.where(rsi < 45, 10, np.where(rsi > 55, -10, 0))))
        signals['rsi_signal'] = pd.Series(rsi_signal, index=df.index).astype(float)

        # 3. MACD（權重20%）
        macd_signal_val = pd.Series(0.0, index=df.index)
        macd_signal_val += np.where(df['MACD'] > df['MACD_Signal'], 10, -10)
        macd_signal_val += np.where(df['MACD_Hist'] > 0, 5, -5)
        # MACD 柱狀圖方向
        macd_hist_dir = df['MACD_Hist'].diff()
        macd_signal_val += np.where(macd_hist_dir > 0, 5, -5)
        signals['macd_signal'] = pd.Series(macd_signal_val, index=df.index).clip(-20, 20)

        # 4. 布林通道（權重15%）
        bb_signal = pd.Series(0.0, index=df.index)
        bb_signal += np.where(df['BB_PctB'] < 0.0, 15, np.where(df['BB_PctB'] > 1.0, -15,
                              np.where(df['BB_PctB'] < 0.2, 8, np.where(df['BB_PctB'] > 0.8, -8, 0))))
        signals['bb_signal'] = pd.Series(bb_signal, index=df.index).astype(float)

        # 5. 成交量確認（權重15%）
        vol_signal = pd.Series(0.0, index=df.index)
        vol_ratio = df['VOL_RATIO']
        price_change = close.pct_change()
        # 量增價升 → 正面信號
        vol_signal += np.where((vol_ratio > 1.5) & (price_change > 0.01), 10, 0)
        # 量增價跌 → 負面信號
        vol_signal += np.where((vol_ratio > 1.5) & (price_change < -0.01), -10, 0)
        # 量縮 → 等待
        vol_signal += np.where(vol_ratio < 0.5, 0, 0)
        signals['vol_signal'] = pd.Series(vol_signal, index=df.index).clip(-15, 15)

        # --- 綜合信號 ---
        signals['total_signal'] = (
            signals['ma_signal'] +
            signals['rsi_signal'] +
            signals['macd_signal'] +
            signals['bb_signal'] +
            signals['vol_signal']
        ).clip(-100, 100)

        # 判斷方向
        signals['direction'] = np.where(signals['total_signal'] > 25, "買入",
                               np.where(signals['total_signal'] < -25, "賣出", "觀望"))

        return signals

    # ==================== 策略三：動量策略 ====================
    def momentum_ranking(self, tickers, start, end, lookback=126, exclude_recent=21):
        """
        截面動量排名：計算所有股票的動量分數並排名
        """
        results = []
        for ticker in tickers:
            try:
                prices = self.de.get_prices(ticker, start, end)
                if prices.empty or len(prices) < lookback + exclude_recent:
                    continue
                close = prices['Close']
                # 動量 = 過去lookback天到exclude_recent天前的回報
                mom = (close.iloc[-(exclude_recent+1)] / close.iloc[-(lookback+exclude_recent+1)]) - 1
                # 波動率調整
                vol = close.pct_change().iloc[-(lookback+exclude_recent):-exclude_recent].std() * np.sqrt(252)
                risk_adj_mom = mom / vol if vol > 0 else 0

                results.append({
                    "ticker": ticker,
                    "raw_momentum": round(float(mom) * 100, 2),
                    "volatility": round(float(vol) * 100, 2),
                    "risk_adj_momentum": round(float(risk_adj_mom), 4),
                    "last_price": round(float(close.iloc[-1]), 2),
                })
            except Exception:
                continue

        df = pd.DataFrame(results)
        if df.empty:
            return df
        df = df.sort_values("risk_adj_momentum", ascending=False).reset_index(drop=True)
        df['rank'] = range(1, len(df) + 1)
        df['quintile'] = pd.qcut(df['risk_adj_momentum'], 5, labels=['Q1_弱', 'Q2', 'Q3', 'Q4', 'Q5_強'])
        return df

    # ==================== 策略四：均值回歸 ====================
    def mean_reversion_signals(self, ticker, start, end, lookback=20):
        """均值回歸策略信號"""
        prices = self.de.get_prices(ticker, start, end)
        if prices.empty:
            return pd.DataFrame()

        close = prices['Close']
        mean = close.rolling(lookback).mean()
        std = close.rolling(lookback).std()
        zscore = (close - mean) / std

        signals = pd.DataFrame(index=prices.index)
        signals['close'] = close
        signals['zscore'] = zscore
        signals['mean'] = mean

        signals['signal'] = np.where(zscore < -2.0, "買入（極度超賣）",
                           np.where(zscore < -1.0, "觀察買入",
                           np.where(zscore > 2.0, "賣出（極度超買）",
                           np.where(zscore > 1.0, "觀察賣出", "中性"))))

        return signals

    # ==================== 綜合策略：六維度評分 ====================
    def comprehensive_analysis(self, ticker, start, end):
        """
        六維度綜合分析：基本面 + 技術面 + 動量 + 價值 + 質量 + 風險
        生成最終的買入/賣出建議與預期回報
        """
        result = {
            "ticker": ticker,
            "analysis_date": pd.Timestamp.now().strftime("%Y-%m-%d"),
        }

        # 1. 多因子評分
        from config import STRATEGY_PARAMS
        mf = self.multifactor_score(ticker, STRATEGY_PARAMS["multifactor"])
        if mf:
            result.update({
                "mf_total": mf["total_score"],
                "mf_quality": mf["quality_score"],
                "mf_value": mf["value_score"],
                "mf_momentum": mf["momentum_score"],
                "f_score": mf["f_score"],
            })

        # 2. 技術面信號
        tech = self.technical_signals(ticker, start, end)
        if not tech.empty:
            latest_signal = tech['total_signal'].iloc[-1]
            avg_signal_5d = tech['total_signal'].iloc[-5:].mean()
            result.update({
                "tech_signal_latest": round(float(latest_signal), 1),
                "tech_signal_5d_avg": round(float(avg_signal_5d), 1),
                "tech_direction": tech['direction'].iloc[-1],
            })

        # 3. 均值回歸
        mr = self.mean_reversion_signals(ticker, start, end)
        if not mr.empty:
            result.update({
                "zscore": round(float(mr['zscore'].iloc[-1]), 2),
                "mr_signal": mr['signal'].iloc[-1],
            })

        # 4. 基本面
        fund = self.de.get_fundamental_scores(ticker)
        if fund:
            result.update({
                "pe": fund.get("pe_trailing"),
                "pb": fund.get("pb"),
                "roe": round(fund["roe"] * 100, 1) if fund.get("roe") else None,
                "debt_equity": fund.get("debt_to_equity"),
                "revenue_growth": round(fund["revenue_growth"] * 100, 1) if fund.get("revenue_growth") else None,
                "profit_margin": round(fund["profit_margin"] * 100, 1) if fund.get("profit_margin") else None,
                "analyst_target": fund.get("target_mean"),
                "analyst_rec": fund.get("recommendation"),
            })

        # 5. 生成預測
        result["prediction"] = self._generate_prediction(result)

        return result

    def _generate_prediction(self, analysis):
        """基於綜合分析生成預測"""
        score = 0
        reasons_bull = []
        reasons_bear = []

        # 多因子評分
        mf_total = analysis.get("mf_total", 50)
        if mf_total >= 70:
            score += 3
            reasons_bull.append(f"多因子評分高({mf_total})")
        elif mf_total >= 55:
            score += 1
        elif mf_total < 35:
            score -= 3
            reasons_bear.append(f"多因子評分低({mf_total})")

        # 技術面
        tech = analysis.get("tech_signal_latest", 0)
        if tech > 30:
            score += 2
            reasons_bull.append(f"技術面看多(信號:{tech})")
        elif tech < -30:
            score -= 2
            reasons_bear.append(f"技術面看空(信號:{tech})")

        # F-Score
        f = analysis.get("f_score", 5)
        if f >= 7:
            score += 2
            reasons_bull.append(f"F-Score高({f}/9)")
        elif f <= 3:
            score -= 2
            reasons_bear.append(f"F-Score低({f}/9)")

        # Z-Score（均值回歸）
        z = analysis.get("zscore", 0)
        if z < -1.5:
            score += 1
            reasons_bull.append(f"均值回歸機會(Z:{z})")
        elif z > 1.5:
            score -= 1
            reasons_bear.append(f"短期過熱(Z:{z})")

        # ROE
        roe = analysis.get("roe")
        if roe and roe > 20:
            score += 1
            reasons_bull.append(f"ROE優秀({roe}%)")

        # 營收成長
        rg = analysis.get("revenue_growth")
        if rg and rg > 15:
            score += 1
            reasons_bull.append(f"營收高成長({rg}%)")
        elif rg and rg < -5:
            score -= 1
            reasons_bear.append(f"營收衰退({rg}%)")

        # 分析師目標價
        target = analysis.get("analyst_target")
        prices_data = self.de.get_prices(analysis["ticker"],
                                         "2025-01-01",
                                         pd.Timestamp.now().strftime("%Y-%m-%d"))
        current_price = None
        if not prices_data.empty:
            current_price = float(prices_data['Close'].iloc[-1])

        upside = None
        if target and current_price and current_price > 0:
            upside = round((target / current_price - 1) * 100, 1)
            if upside > 20:
                score += 1
                reasons_bull.append(f"分析師目標上行空間{upside}%")

        # 判定
        if score >= 5:
            direction = "強烈看漲"
            expected_return = "15-25%"
            timeframe = "3-6個月"
            confidence = "高"
        elif score >= 3:
            direction = "看漲"
            expected_return = "8-15%"
            timeframe = "3-6個月"
            confidence = "中高"
        elif score >= 1:
            direction = "輕度看漲"
            expected_return = "3-8%"
            timeframe = "1-3個月"
            confidence = "中"
        elif score >= -1:
            direction = "中性/觀望"
            expected_return = "-3%~3%"
            timeframe = "—"
            confidence = "低"
        elif score >= -3:
            direction = "看跌"
            expected_return = "-5%~-15%"
            timeframe = "1-3個月"
            confidence = "中"
        else:
            direction = "強烈看跌"
            expected_return = "-15%以下"
            timeframe = "1-3個月"
            confidence = "高"

        return {
            "direction": direction,
            "score": score,
            "expected_return": expected_return,
            "timeframe": timeframe,
            "confidence": confidence,
            "current_price": current_price,
            "analyst_target": target,
            "upside_pct": upside,
            "bull_reasons": reasons_bull,
            "bear_reasons": reasons_bear,
        }
