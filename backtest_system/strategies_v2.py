"""
策略引擎 v2 — 全面優化版
修正內容：
1. 整合 Finnhub 分析師評級 + 內部人交易
2. 整合 FMP DCF估值
3. 行業中性化（金融用PB、科技用PS、醫藥加管線考量）
4. 估值過熱懲罰（PE>50 扣分）
5. 技術面信號閾值重新校準
6. FRED 宏觀環境開關（VIX/利率/殖利率曲線）
7. 動態因子權重（趨勢市加重動量，震盪市加重價值）
"""

import pandas as pd
import numpy as np
import requests
from datetime import datetime, timedelta
from data_engine import DataEngine

# ==================== 行業分類 ====================
SECTOR_MAP = {
    # 金融股：用PB估值，不看PE
    "Financial Services": "financial",
    "Banking": "financial",
    "Insurance": "financial",
    "Financial": "financial",
    # 科技股：用PS和EV/EBITDA，容忍高PE
    "Technology": "tech",
    "Semiconductors": "tech",
    "Software": "tech",
    "Communication Services": "tech",
    "Media": "tech",
    # 醫藥股：加入成長性和管線權重
    "Healthcare": "healthcare",
    "Biotechnology": "healthcare",
    # 消費股：看穩定性
    "Consumer Cyclical": "consumer",
    "Consumer Defensive": "consumer_staple",
    # 能源/工業：看周期
    "Energy": "cyclical",
    "Industrials": "cyclical",
    "Basic Materials": "cyclical",
    # 其他
    "Utilities": "defensive",
    "Real Estate": "defensive",
}


class StrategyV2:
    """策略引擎 v2 — 全面優化"""

    def __init__(self, data_engine: DataEngine, finnhub=None, fmp_key=None, fred_key=None):
        self.de = data_engine
        self.fh = finnhub
        self.fmp_key = fmp_key
        self.fred_key = fred_key
        self._macro_cache = {}

    # ==================== 宏觀環境判斷 ====================
    def get_macro_regime(self):
        """
        判斷當前宏觀環境：Risk-On / Neutral / Risk-Off
        基於 VIX、殖利率曲線、金融壓力指數
        """
        cache_key = "macro_regime"
        if cache_key in self._macro_cache:
            return self._macro_cache[cache_key]

        regime = {"state": "neutral", "vix": None, "yield_curve": None, "score": 0}

        try:
            import yfinance as yf
            # VIX
            vix_data = yf.download("^VIX", period="5d", progress=False)
            if not vix_data.empty:
                if isinstance(vix_data.columns, pd.MultiIndex):
                    vix_data.columns = vix_data.columns.get_level_values(0)
                vix = float(vix_data['Close'].iloc[-1])
                regime["vix"] = round(vix, 1)
                if vix < 15:
                    regime["score"] += 2  # 極低波動 → Risk-On
                elif vix < 20:
                    regime["score"] += 1
                elif vix > 30:
                    regime["score"] -= 2  # 恐慌 → Risk-Off
                elif vix > 25:
                    regime["score"] -= 1

            # 殖利率曲線（用 FRED 或 yfinance 代理）
            tnx = yf.download("^TNX", period="5d", progress=False)  # 10Y
            twoy = yf.download("^IRX", period="5d", progress=False)  # 3M（近似2Y）
            if not tnx.empty and not twoy.empty:
                if isinstance(tnx.columns, pd.MultiIndex):
                    tnx.columns = tnx.columns.get_level_values(0)
                if isinstance(twoy.columns, pd.MultiIndex):
                    twoy.columns = twoy.columns.get_level_values(0)
                spread = float(tnx['Close'].iloc[-1]) - float(twoy['Close'].iloc[-1])
                regime["yield_curve"] = round(spread, 2)
                if spread > 1.0:
                    regime["score"] += 1  # 正常陡峭
                elif spread < 0:
                    regime["score"] -= 2  # 倒掛 → 衰退風險

            # 市場趨勢（SPY 50MA vs 200MA）
            spy = yf.download("SPY", period="1y", progress=False)
            if not spy.empty:
                if isinstance(spy.columns, pd.MultiIndex):
                    spy.columns = spy.columns.get_level_values(0)
                spy_close = spy['Close']
                sma50 = spy_close.rolling(50).mean().iloc[-1]
                sma200 = spy_close.rolling(200).mean().iloc[-1]
                if sma50 > sma200:
                    regime["score"] += 1  # 趨勢向上
                else:
                    regime["score"] -= 1

        except Exception:
            pass

        # 判斷環境
        if regime["score"] >= 3:
            regime["state"] = "risk_on"
        elif regime["score"] <= -2:
            regime["state"] = "risk_off"
        else:
            regime["state"] = "neutral"

        self._macro_cache[cache_key] = regime
        return regime

    def get_dynamic_weights(self, macro_regime):
        """根據宏觀環境動態調整因子權重"""
        state = macro_regime.get("state", "neutral")

        if state == "risk_on":
            # 趨勢市：加重動量和成長
            return {
                "quality": 0.20, "value": 0.15, "momentum": 0.30,
                "analyst": 0.15, "insider": 0.05, "dcf": 0.15,
            }
        elif state == "risk_off":
            # 避險市：加重質量和價值
            return {
                "quality": 0.30, "value": 0.25, "momentum": 0.10,
                "analyst": 0.15, "insider": 0.10, "dcf": 0.10,
            }
        else:
            # 中性：均衡
            return {
                "quality": 0.25, "value": 0.20, "momentum": 0.20,
                "analyst": 0.15, "insider": 0.05, "dcf": 0.15,
            }

    # ==================== 行業感知估值 ====================
    def sector_adjusted_value_score(self, info, sector_type):
        """行業中性化的估值評分"""
        score = 0
        max_score = 5
        reasons = []

        pe = info.get("forwardPE") or info.get("trailingPE")
        pb = info.get("priceToBook")
        ps = info.get("priceToSalesTrailing12Months")
        peg = info.get("pegRatio")
        dy = info.get("dividendYield")

        if sector_type == "financial":
            # 金融股：主看 PB，不看 PE
            if pb is not None:
                if 0 < pb < 1.2:
                    score += 3
                    reasons.append(f"PB低估({pb:.1f})")
                elif 0 < pb < 2.0:
                    score += 1.5
                elif pb > 3.0:
                    score -= 1
                    reasons.append(f"PB偏高({pb:.1f})")
            if dy and dy > 0.03:
                score += 1
                reasons.append(f"股息率{dy*100:.1f}%")
            if pe and pe > 0:
                score += 0.5  # 金融股PE只做參考

        elif sector_type == "tech":
            # 科技股：看 PS、PEG，容忍高 PE
            if peg is not None:
                if 0 < peg < 1.0:
                    score += 2.5
                    reasons.append(f"PEG極低({peg:.2f})")
                elif 0 < peg < 1.5:
                    score += 1.5
                elif 0 < peg < 2.0:
                    score += 0.5
            if ps is not None:
                if 0 < ps < 5:
                    score += 1.5
                elif 0 < ps < 10:
                    score += 0.5
                elif ps > 20:
                    score -= 1
                    reasons.append(f"PS過高({ps:.1f})")
            # 科技股PE>60才算過熱
            if pe and pe > 60:
                score -= 1.5
                reasons.append(f"PE過高({pe:.0f})")

        elif sector_type == "healthcare":
            # 醫藥股：看PEG和成長
            if peg is not None and 0 < peg < 2:
                score += 2
            eg = info.get("earningsGrowth")
            if eg and eg > 0.2:
                score += 2
                reasons.append(f"盈餘高成長({eg*100:.0f}%)")
            if pe and pe > 80:
                score -= 1
                reasons.append(f"PE過高({pe:.0f})")

        else:
            # 傳統估值：PE + PB + PEG
            if pe is not None:
                if 0 < pe < 12:
                    score += 2
                elif 0 < pe < 20:
                    score += 1
                elif pe > 40:
                    score -= 1
                    reasons.append(f"PE偏高({pe:.0f})")
                if pe > 50:
                    score -= 1.5  # 估值過熱懲罰
                    reasons.append(f"估值過熱警告(PE={pe:.0f})")
            if pb is not None:
                if 0 < pb < 2:
                    score += 1.5
                elif 0 < pb < 4:
                    score += 0.5
            if peg is not None and 0 < peg < 1.5:
                score += 1
            if dy and dy > 0.02:
                score += 0.5

        return min(max(score / max_score, 0), 1.0) * 100, reasons

    # ==================== Finnhub 分析師+內部人評分 ====================
    def analyst_score(self, ticker):
        """分析師評級評分（0-100）"""
        if not self.fh:
            return 50, []

        reasons = []
        score = 50  # 預設中性

        ratings = self.fh.get_analyst_ratings(ticker)
        if ratings:
            bull_pct = ratings.get("bull_pct", 50)
            if bull_pct > 80:
                score = 90
                reasons.append(f"分析師{bull_pct}%看多({ratings['consensus']})")
            elif bull_pct > 65:
                score = 75
                reasons.append(f"分析師{bull_pct}%看多")
            elif bull_pct > 50:
                score = 60
            elif bull_pct < 30:
                score = 20
                reasons.append(f"分析師僅{bull_pct}%看多")
            elif bull_pct < 45:
                score = 35

        return score, reasons

    def insider_score(self, ticker):
        """內部人交易評分（0-100）"""
        if not self.fh:
            return 50, []

        reasons = []
        score = 50

        insider = self.fh.get_insider_transactions(ticker)
        if insider:
            signal = insider.get("net_signal", "")
            if "強烈看多" in signal:
                score = 90
                reasons.append(f"內部人大量買入(買{insider['buys']}/賣{insider['sells']})")
            elif "偏多" in signal:
                score = 70
                reasons.append(f"內部人淨買入")
            elif "看空" in signal:
                score = 25
                reasons.append(f"內部人大量賣出(買{insider['buys']}/賣{insider['sells']})")
            elif "偏空" in signal:
                score = 35

        return score, reasons

    # ==================== FMP DCF 估值評分 ====================
    def dcf_score(self, ticker):
        """DCF估值評分：當前價vs DCF公允價值"""
        if not self.fmp_key:
            return 50, []

        reasons = []
        score = 50

        try:
            base = "https://financialmodelingprep.com/stable"
            r = requests.get(f"{base}/discounted-cash-flow?symbol={ticker}&apikey={self.fmp_key}", timeout=10)
            data = r.json()
            if data and isinstance(data, list) and len(data) > 0:
                dcf_value = data[0].get("dcf")
                stock_price = data[0].get("stockPrice")

                if dcf_value and stock_price and stock_price > 0:
                    upside = (dcf_value / stock_price - 1) * 100

                    if upside > 40:
                        score = 95
                        reasons.append(f"DCF嚴重低估({upside:.0f}%上行空間)")
                    elif upside > 20:
                        score = 80
                        reasons.append(f"DCF低估({upside:.0f}%上行)")
                    elif upside > 5:
                        score = 60
                    elif upside > -10:
                        score = 45
                    elif upside > -25:
                        score = 30
                        reasons.append(f"DCF顯示高估({upside:.0f}%)")
                    else:
                        score = 15
                        reasons.append(f"DCF嚴重高估({upside:.0f}%)")
        except Exception:
            pass

        return score, reasons

    # ==================== 技術面信號 v2 ====================
    def technical_signal_v2(self, ticker, start, end):
        """
        技術面信號 v2：修正閾值，增加趨勢強度分級
        回傳：-100 到 +100 的信號分數
        """
        prices = self.de.get_prices(ticker, start, end)
        if prices.empty or len(prices) < 50:
            return 0, []

        df = self.de.calc_all_technicals(prices)
        close = df['Close']
        reasons = []
        score = 0

        # 1. 趨勢判斷（權重最大）
        sma50 = df['SMA_50'].iloc[-1] if 'SMA_50' in df.columns else None
        sma200 = df['SMA_200'].iloc[-1] if 'SMA_200' in df.columns else None
        current = float(close.iloc[-1])

        if sma50 and sma200:
            if current > sma50 > sma200:
                score += 25
                reasons.append("強勢趨勢（價格>50MA>200MA）")
            elif current > sma200 and current > sma50:
                score += 15
            elif current > sma200:
                score += 5
            elif current < sma50 < sma200:
                score -= 25
                reasons.append("弱勢趨勢（價格<50MA<200MA）")
            elif current < sma200:
                score -= 15

        # 2. RSI（避免過度依賴）
        rsi = df['RSI_14'].iloc[-1] if 'RSI_14' in df.columns else 50
        if rsi < 25:
            score += 15
            reasons.append(f"RSI極度超賣({rsi:.0f})")
        elif rsi < 35:
            score += 8
        elif rsi > 75:
            score -= 15
            reasons.append(f"RSI極度超買({rsi:.0f})")
        elif rsi > 65:
            score -= 5

        # 3. MACD（看方向變化而非絕對值）
        if 'MACD_Hist' in df.columns:
            hist = df['MACD_Hist']
            hist_now = float(hist.iloc[-1])
            hist_prev = float(hist.iloc[-2]) if len(hist) > 1 else 0
            # 柱狀圖由負轉正 = 買入信號
            if hist_now > 0 and hist_prev < 0:
                score += 15
                reasons.append("MACD金叉確認")
            elif hist_now < 0 and hist_prev > 0:
                score -= 15
                reasons.append("MACD死叉確認")
            elif hist_now > 0 and hist_now > hist_prev:
                score += 5  # 動能增強
            elif hist_now < 0 and hist_now < hist_prev:
                score -= 5  # 動能減弱

        # 4. 布林通道（只在極端時觸發）
        if 'BB_PctB' in df.columns:
            pctb = float(df['BB_PctB'].iloc[-1])
            if pctb < -0.1:
                score += 10
                reasons.append("跌破布林下軌")
            elif pctb > 1.1:
                score -= 10
                reasons.append("突破布林上軌")

        # 5. 成交量確認
        if 'VOL_RATIO' in df.columns:
            vol_ratio = float(df['VOL_RATIO'].iloc[-1])
            price_chg = float(close.pct_change().iloc[-1])
            if vol_ratio > 2.0 and price_chg > 0.02:
                score += 10
                reasons.append(f"放量上漲(量比{vol_ratio:.1f})")
            elif vol_ratio > 2.0 and price_chg < -0.02:
                score -= 10
                reasons.append(f"放量下跌(量比{vol_ratio:.1f})")

        # 6. 動量一致性（多時間框架確認）
        mom_scores = 0
        for period in [21, 63, 126]:
            col = f'MOM_{period}'
            if col in df.columns:
                mom = float(df[col].iloc[-1])
                if mom > 0:
                    mom_scores += 1
                elif mom < 0:
                    mom_scores -= 1
        if mom_scores == 3:
            score += 10
            reasons.append("多時間框架動量一致看多")
        elif mom_scores == -3:
            score -= 10
            reasons.append("多時間框架動量一致看空")

        return max(min(score, 100), -100), reasons

    # ==================== 綜合分析 v2 ====================
    def comprehensive_analysis_v2(self, ticker, start, end):
        """
        v2 綜合分析：整合所有數據源，行業中性化，動態權重
        """
        result = {"ticker": ticker, "analysis_date": datetime.now().strftime("%Y-%m-%d")}

        # 0. 宏觀環境
        macro = self.get_macro_regime()
        result["macro_regime"] = macro["state"]
        result["vix"] = macro.get("vix")

        # 動態權重
        weights = self.get_dynamic_weights(macro)

        # 1. 基本面數據
        info = self.de.get_stock_info(ticker)
        sector = info.get("sector", "") if info else ""
        industry = info.get("industry", "") if info else ""
        sector_type = SECTOR_MAP.get(sector, SECTOR_MAP.get(industry, "other"))
        result["sector"] = sector
        result["sector_type"] = sector_type

        all_reasons_bull = []
        all_reasons_bear = []

        # 2. 質量因子
        q_score = 0
        if info:
            roe = info.get("returnOnEquity")
            gm = info.get("grossMargins")
            dte = info.get("debtToEquity")
            fcf = info.get("freeCashflow")
            om = info.get("operatingMargins")

            pts = 0
            mx = 5
            if roe and roe > 0.20:
                pts += 2
            elif roe and roe > 0.10:
                pts += 1
            if gm and gm > 0.40:
                pts += 1
            elif gm and gm > 0.25:
                pts += 0.5
            if dte is not None and dte < 80:
                pts += 0.5
            if fcf and fcf > 0:
                pts += 0.5
            if om and om > 0.20:
                pts += 0.5

            q_score = min(pts / mx, 1.0) * 100
        result["quality_score"] = round(q_score, 1)

        # 3. 價值因子（行業中性化）
        v_score, v_reasons = self.sector_adjusted_value_score(info, sector_type) if info else (50, [])
        result["value_score"] = round(v_score, 1)
        all_reasons_bull.extend([r for r in v_reasons if "低估" in r or "低" in r.split("(")[0]])
        all_reasons_bear.extend([r for r in v_reasons if "高" in r or "過熱" in r])

        # 4. 動量因子
        m_score = 50
        try:
            prices = self.de.get_prices(ticker, start, end)
            if not prices.empty and len(prices) > 126:
                close = prices['Close']
                mom_6m = float((close.iloc[-22] / close.iloc[-148]) - 1) if len(close) > 148 else 0
                sma200 = close.rolling(200).mean()
                above_200 = float(close.iloc[-1]) > float(sma200.iloc[-1]) if len(sma200.dropna()) > 0 else False

                m_pts = 0
                if mom_6m > 0.20:
                    m_pts += 3
                    all_reasons_bull.append(f"6月動量強勁(+{mom_6m*100:.0f}%)")
                elif mom_6m > 0.05:
                    m_pts += 1.5
                elif mom_6m < -0.15:
                    m_pts -= 2
                    all_reasons_bear.append(f"6月動量疲弱({mom_6m*100:.0f}%)")

                if above_200:
                    m_pts += 1.5
                else:
                    m_pts -= 1

                m_score = min(max((m_pts + 3) / 6, 0), 1.0) * 100
        except Exception:
            pass
        result["momentum_score"] = round(m_score, 1)

        # 5. 分析師評級（Finnhub）
        a_score, a_reasons = self.analyst_score(ticker)
        result["analyst_score"] = round(a_score, 1)
        all_reasons_bull.extend([r for r in a_reasons if "看多" in r])
        all_reasons_bear.extend([r for r in a_reasons if "僅" in r])

        # 6. 內部人交易（Finnhub）
        i_score, i_reasons = self.insider_score(ticker)
        result["insider_score"] = round(i_score, 1)
        all_reasons_bull.extend([r for r in i_reasons if "買入" in r])
        all_reasons_bear.extend([r for r in i_reasons if "賣出" in r])

        # 7. DCF估值（FMP）
        d_score, d_reasons = self.dcf_score(ticker)
        result["dcf_score"] = round(d_score, 1)
        all_reasons_bull.extend([r for r in d_reasons if "低估" in r])
        all_reasons_bear.extend([r for r in d_reasons if "高估" in r])

        # 8. 技術面
        t_score, t_reasons = self.technical_signal_v2(ticker, start, end)
        # 轉換到 0-100 範圍
        tech_normalized = (t_score + 100) / 2
        result["tech_score"] = round(tech_normalized, 1)
        all_reasons_bull.extend([r for r in t_reasons if any(k in r for k in ["強勢", "超賣", "金叉", "放量上漲", "看多"])])
        all_reasons_bear.extend([r for r in t_reasons if any(k in r for k in ["弱勢", "超買", "死叉", "放量下跌", "看空"])])

        # ==================== 加權綜合評分 ====================
        total_score = (
            q_score * weights["quality"] +
            v_score * weights["value"] +
            m_score * weights["momentum"] +
            a_score * weights["analyst"] +
            i_score * weights["insider"] +
            d_score * weights["dcf"]
        )

        # 技術面作為確認/否決（不直接加入權重，作為修正項）
        if t_score > 30:
            total_score *= 1.10  # 技術面確認，加10%
        elif t_score < -30:
            total_score *= 0.85  # 技術面否決，減15%

        # F-Score 加分
        f_data = self.de.calc_piotroski_f_score(ticker)
        f_score = f_data.get("score", 5)
        result["f_score"] = f_score
        if f_score >= 8:
            total_score *= 1.08
            all_reasons_bull.append(f"F-Score極高({f_score}/9)")
        elif f_score <= 2:
            total_score *= 0.88
            all_reasons_bear.append(f"F-Score極低({f_score}/9)")

        total_score = max(min(total_score, 100), 0)
        result["total_score"] = round(total_score, 1)

        # ==================== 生成預測 ====================
        result["prediction"] = self._generate_prediction_v2(
            total_score, t_score, all_reasons_bull, all_reasons_bear,
            ticker, macro
        )

        return result

    def _generate_prediction_v2(self, total_score, tech_score,
                                 reasons_bull, reasons_bear, ticker, macro):
        """v2 預測生成"""
        # 根據環境調整閾值
        state = macro.get("state", "neutral")
        if state == "risk_off":
            buy_threshold = 68  # 避險時提高買入門檻
            strong_buy_threshold = 78
        elif state == "risk_on":
            buy_threshold = 55  # 趨勢好時降低門檻
            strong_buy_threshold = 70
        else:
            buy_threshold = 60
            strong_buy_threshold = 75

        if total_score >= strong_buy_threshold:
            direction = "強烈看漲"
            expected_return = "15-25%"
            timeframe = "3-6個月"
            confidence = "高"
        elif total_score >= buy_threshold:
            direction = "看漲"
            expected_return = "8-15%"
            timeframe = "3-6個月"
            confidence = "中高"
        elif total_score >= 48:
            direction = "輕度看漲"
            expected_return = "3-8%"
            timeframe = "1-3個月"
            confidence = "中"
        elif total_score >= 38:
            direction = "中性/觀望"
            expected_return = "-3%~3%"
            timeframe = "—"
            confidence = "低"
        elif total_score >= 28:
            direction = "看跌"
            expected_return = "-5%~-15%"
            timeframe = "1-3個月"
            confidence = "中"
        else:
            direction = "強烈看跌"
            expected_return = "-15%以下"
            timeframe = "1-3個月"
            confidence = "高"

        # 技術面可以否決基本面信號
        if direction in ["強烈看漲", "看漲"] and tech_score < -40:
            direction = "觀望（基本面好但技術面弱）"
            confidence = "低"
            reasons_bear.append("技術面強烈否決，建議等待技術面改善再入場")

        if direction in ["強烈看跌", "看跌"] and tech_score > 40:
            direction = "觀望（基本面弱但技術面強）"
            confidence = "低"
            reasons_bull.append("技術面支撐，短期可能反彈")

        return {
            "direction": direction,
            "total_score": round(total_score, 1),
            "tech_score": tech_score,
            "expected_return": expected_return,
            "timeframe": timeframe,
            "confidence": confidence,
            "macro_regime": macro.get("state"),
            "bull_reasons": reasons_bull[:5],
            "bear_reasons": reasons_bear[:5],
        }

    @staticmethod
    def get_recommendation(score):
        if score >= 75: return "強烈買入"
        elif score >= 60: return "買入"
        elif score >= 48: return "輕度看漲"
        elif score >= 38: return "觀望"
        elif score >= 28: return "減持"
        else: return "賣出"
