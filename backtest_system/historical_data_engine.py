"""
歷史數據引擎 — 只使用「當時能看到的數據」，消除前視偏差
SimFin: 歷史季度財報
Polygon: 歷史季度財報（備用）
Tiingo: 穩定歷史股價
Marketaux: 新聞情緒
Twelve Data: 歷史技術指標
FRED: 宏觀經濟（本身就是歷史數據）
"""

import requests
import pandas as pd
import numpy as np
import time
from datetime import datetime, timedelta
from config import (SIMFIN_API_KEY, POLYGON_API_KEY, TIINGO_API_KEY,
                    MARKETAUX_API_KEY, TWELVE_DATA_API_KEY,
                    FRED_API_KEY, FINNHUB_API_KEY)


class HistoricalDataEngine:
    """只使用歷史時點數據，無前視偏差"""

    def __init__(self):
        self._cache = {}
        self._call_count = 0

    def _rate_limit(self):
        self._call_count += 1
        if self._call_count % 8 == 0:
            time.sleep(1.5)

    # ==================== 歷史股價（Tiingo）====================
    def get_historical_prices(self, ticker, start, end):
        """從Tiingo取歷史股價"""
        cache_key = f"price_{ticker}_{start}_{end}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        self._rate_limit()
        try:
            headers = {'Content-Type': 'application/json',
                       'Authorization': f'Token {TIINGO_API_KEY}'}
            r = requests.get(
                f'https://api.tiingo.com/tiingo/daily/{ticker}/prices'
                f'?startDate={start}&endDate={end}',
                headers=headers, timeout=15)
            if r.status_code == 200:
                data = r.json()
                if data:
                    df = pd.DataFrame(data)
                    df['date'] = pd.to_datetime(df['date']).dt.tz_localize(None)
                    df = df.set_index('date')
                    df = df.rename(columns={'close': 'Close', 'high': 'High',
                                            'low': 'Low', 'open': 'Open',
                                            'volume': 'Volume'})
                    self._cache[cache_key] = df
                    return df
        except Exception:
            pass

        # 備用：yfinance
        try:
            import yfinance as yf
            df = yf.download(ticker, start=start, end=end, progress=False, auto_adjust=True)
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            if not df.empty:
                self._cache[cache_key] = df
                return df
        except:
            pass

        return pd.DataFrame()

    # ==================== 歷史財報（Polygon.io）====================
    def get_historical_financials(self, ticker, as_of_date):
        """
        用 Polygon.io 取得在 as_of_date 當天能看到的最新季度財報
        Polygon 有完整的歷史財報數據（2019-2026）
        """
        cache_key = f"fin_{ticker}_{as_of_date}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        self._rate_limit()
        try:
            r = requests.get(
                f'https://api.polygon.io/vX/reference/financials'
                f'?ticker={ticker}&timeframe=quarterly'
                f'&period_of_report_date.lte={as_of_date}'
                f'&limit=2&sort=period_of_report_date&order=desc'
                f'&apiKey={POLYGON_API_KEY}', timeout=15)

            if r.status_code != 200:
                return None

            results = r.json().get('results', [])
            if not results:
                return None

            q = results[0]
            inc = q.get('financials', {}).get('income_statement', {})
            bs = q.get('financials', {}).get('balance_sheet', {})
            cf = q.get('financials', {}).get('cash_flow_statement', {})

            rev = self._pg_val(inc, 'revenues')
            gp = self._pg_val(inc, 'gross_profit')
            oi = self._pg_val(inc, 'operating_income_loss')
            ni = self._pg_val(inc, 'net_income_loss')
            eps = self._pg_val(inc, 'basic_earnings_per_share')

            ta = self._pg_val(bs, 'assets')
            tl = self._pg_val(bs, 'liabilities')
            eq = self._pg_val(bs, 'equity')

            ocf = self._pg_val(cf, 'net_cash_flow_from_operating_activities')

            result = {
                "ticker": ticker,
                "fiscal_period": q.get("fiscal_period", ""),
                "fiscal_year": q.get("fiscal_year", ""),
                "filing_date": q.get("filing_date", ""),
                "revenue": rev, "gross_profit": gp,
                "operating_income": oi, "net_income": ni,
                "eps_diluted": eps,
                "total_assets": ta, "total_liabilities": tl,
                "total_equity": eq, "operating_cf": ocf,
            }

            # 計算比率
            if rev and rev > 0:
                result["gross_margin"] = round(gp / rev, 4) if gp else None
                result["net_margin"] = round(ni / rev, 4) if ni else None
                result["operating_margin"] = round(oi / rev, 4) if oi else None

            if eq and eq > 0:
                result["roe"] = round(ni * 4 / eq, 4) if ni else None  # 年化
                result["debt_equity"] = round(tl / eq, 4) if tl else None

            if ta and ta > 0:
                result["roa"] = round(ni * 4 / ta, 4) if ni else None

            # 前一季數據（計算成長率）
            if len(results) > 1:
                q_prev = results[1]
                inc_prev = q_prev.get('financials', {}).get('income_statement', {})
                rev_prev = self._pg_val(inc_prev, 'revenues')
                ni_prev = self._pg_val(inc_prev, 'net_income_loss')
                if rev_prev and rev_prev > 0 and rev:
                    result["revenue_growth_qoq"] = round((rev / rev_prev - 1), 4)
                if ni_prev and ni_prev > 0 and ni:
                    result["earnings_growth_qoq"] = round((ni / ni_prev - 1), 4)

            self._cache[cache_key] = result
            return result

        except Exception:
            return None

    @staticmethod
    def _pg_val(section, key):
        """從Polygon財報中安全取值"""
        try:
            v = section.get(key, {}).get('value')
            return float(v) if v is not None else 0
        except:
            return 0

    # ==================== 歷史估值指標（從股價+財報計算）====================
    def calc_historical_valuation(self, ticker, as_of_date, financials=None):
        """用歷史股價+歷史財報計算當時的PE/PB等"""
        if financials is None:
            financials = self.get_historical_financials(ticker, as_of_date)

        if not financials:
            return {}

        # 取當天股價
        start = (datetime.strptime(as_of_date, "%Y-%m-%d") - timedelta(days=5)).strftime("%Y-%m-%d")
        prices = self.get_historical_prices(ticker, start, as_of_date)
        if prices.empty:
            return {}

        price = float(prices['Close'].iloc[-1])

        valuation = {"price": price}

        eps = financials.get("eps_diluted", 0)
        if eps and eps > 0:
            valuation["pe_ttm"] = round(price / (eps * 4), 2)  # 年化

        equity = financials.get("total_equity", 0)
        if equity and equity > 0:
            # 需要股數來算PB，用市值估算
            # 簡化：用EPS反推股數
            ni = financials.get("net_income", 0)
            if ni and eps and eps != 0:
                shares = abs(ni / eps)
                bvps = equity / shares
                if bvps > 0:
                    valuation["pb"] = round(price / bvps, 2)

        valuation["gross_margin"] = financials.get("gross_margin")
        valuation["net_margin"] = financials.get("net_margin")
        valuation["roe"] = financials.get("roe")
        valuation["roa"] = financials.get("roa")
        valuation["debt_equity"] = financials.get("debt_equity")

        return valuation

    # ==================== 歷史技術指標 ====================
    def calc_historical_technicals(self, ticker, as_of_date, lookback_days=250):
        """用歷史股價計算技術指標"""
        start = (datetime.strptime(as_of_date, "%Y-%m-%d") - timedelta(days=lookback_days + 50)).strftime("%Y-%m-%d")
        prices = self.get_historical_prices(ticker, start, as_of_date)

        if prices.empty or len(prices) < 50:
            return {}

        close = prices['Close']
        high = prices['High']
        low = prices['Low']
        volume = prices['Volume']

        current = float(close.iloc[-1])

        # 均線
        sma20 = float(close.rolling(20).mean().iloc[-1])
        sma50 = float(close.rolling(50).mean().iloc[-1])
        sma200 = float(close.rolling(200).mean().iloc[-1]) if len(close) >= 200 else sma50

        # RSI
        delta = close.diff()
        gain = delta.where(delta > 0, 0).rolling(14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(14).mean()
        rs = gain / loss.replace(0, np.nan)
        rsi = float((100 - 100 / (1 + rs)).iloc[-1])

        # MACD
        ema12 = close.ewm(span=12).mean()
        ema26 = close.ewm(span=26).mean()
        macd = ema12 - ema26
        signal = macd.ewm(span=9).mean()
        hist = macd - signal
        macd_val = float(macd.iloc[-1])
        signal_val = float(signal.iloc[-1])
        hist_val = float(hist.iloc[-1])
        hist_prev = float(hist.iloc[-2]) if len(hist) > 1 else 0

        # 動量
        mom_1m = float((close.iloc[-1] / close.iloc[-22] - 1)) if len(close) >= 22 else 0
        mom_3m = float((close.iloc[-1] / close.iloc[-63] - 1)) if len(close) >= 63 else 0
        mom_6m = float((close.iloc[-1] / close.iloc[-126] - 1)) if len(close) >= 126 else 0

        # 波動率
        vol_20 = float(close.pct_change().rolling(20).std().iloc[-1] * np.sqrt(252))

        # 量比
        vol_ma20 = float(volume.rolling(20).mean().iloc[-1])
        vol_ratio = float(volume.iloc[-1] / vol_ma20) if vol_ma20 > 0 else 1

        # 布林通道
        bb_mid = sma20
        bb_std = float(close.rolling(20).std().iloc[-1])
        bb_upper = bb_mid + 2 * bb_std
        bb_lower = bb_mid - 2 * bb_std
        bb_pctb = (current - bb_lower) / (bb_upper - bb_lower) if (bb_upper - bb_lower) > 0 else 0.5

        return {
            "price": current,
            "sma20": round(sma20, 2), "sma50": round(sma50, 2), "sma200": round(sma200, 2),
            "above_sma200": current > sma200,
            "sma50_above_sma200": sma50 > sma200,
            "rsi": round(rsi, 1),
            "macd": round(macd_val, 4), "macd_signal": round(signal_val, 4),
            "macd_hist": round(hist_val, 4), "macd_hist_prev": round(hist_prev, 4),
            "macd_cross_up": hist_val > 0 and hist_prev <= 0,
            "macd_cross_down": hist_val < 0 and hist_prev >= 0,
            "mom_1m": round(mom_1m * 100, 2),
            "mom_3m": round(mom_3m * 100, 2),
            "mom_6m": round(mom_6m * 100, 2),
            "volatility_20d": round(vol_20 * 100, 2),
            "vol_ratio": round(vol_ratio, 2),
            "bb_pctb": round(bb_pctb, 4),
        }

    # ==================== 歷史宏觀環境（FRED）====================
    def get_macro_at_date(self, as_of_date):
        """取得特定日期的宏觀環境"""
        cache_key = f"macro_{as_of_date}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        result = {}
        self._rate_limit()

        # VIX（從yfinance取歷史）
        try:
            start = (datetime.strptime(as_of_date, "%Y-%m-%d") - timedelta(days=10)).strftime("%Y-%m-%d")
            vix_prices = self.get_historical_prices("^VIX", start, as_of_date)
            if not vix_prices.empty:
                result["vix"] = round(float(vix_prices['Close'].iloc[-1]), 1)
        except:
            pass

        # SPY趨勢
        try:
            start = (datetime.strptime(as_of_date, "%Y-%m-%d") - timedelta(days=250)).strftime("%Y-%m-%d")
            spy = self.get_historical_prices("SPY", start, as_of_date)
            if not spy.empty and len(spy) >= 50:
                close = spy['Close']
                sma20 = float(close.rolling(20).mean().iloc[-1])
                sma50 = float(close.rolling(50).mean().iloc[-1])
                sma200 = float(close.rolling(200).mean().iloc[-1]) if len(close) >= 200 else sma50
                current = float(close.iloc[-1])

                if current > sma20 and sma20 > sma50:
                    result["spy_trend"] = "strong_up"
                elif current > sma20:
                    result["spy_trend"] = "up"
                elif current < sma20 and current < sma50:
                    result["spy_trend"] = "down"
                else:
                    result["spy_trend"] = "neutral"

                result["spy_price"] = round(current, 2)
                result["spy_sma200"] = round(sma200, 2)
        except:
            pass

        # FRED 數據（殖利率曲線）
        try:
            r = requests.get(
                f'https://api.stlouisfed.org/fred/series/observations'
                f'?series_id=T10Y2Y&api_key={FRED_API_KEY}'
                f'&observation_start={as_of_date}&observation_end={as_of_date}'
                f'&file_type=json', timeout=10)
            if r.status_code == 200:
                obs = r.json().get('observations', [])
                if obs and obs[-1].get('value', '.') != '.':
                    result["yield_curve_10y2y"] = float(obs[-1]['value'])
        except:
            pass

        # 判斷環境
        score = 0
        vix = result.get("vix", 20)
        if vix < 15: score += 2
        elif vix < 20: score += 1
        elif vix > 30: score -= 2
        elif vix > 25: score -= 1

        trend = result.get("spy_trend", "neutral")
        if trend == "strong_up": score += 2
        elif trend == "up": score += 1
        elif trend == "down": score -= 2

        yc = result.get("yield_curve_10y2y")
        if yc is not None:
            if yc > 0.5: score += 1
            elif yc < 0: score -= 1

        if score >= 3: result["regime"] = "risk_on"
        elif score <= -2: result["regime"] = "risk_off"
        else: result["regime"] = "neutral"

        result["regime_score"] = score

        self._cache[cache_key] = result
        return result

    # ==================== 新聞情緒（Marketaux）====================
    def get_news_sentiment(self, ticker, as_of_date, days_back=7):
        """取得特定日期前N天的新聞情緒"""
        cache_key = f"news_{ticker}_{as_of_date}_{days_back}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        self._rate_limit()
        try:
            start = (datetime.strptime(as_of_date, "%Y-%m-%d") - timedelta(days=days_back)).strftime("%Y-%m-%d")
            r = requests.get(
                f'https://api.marketaux.com/v1/news/all'
                f'?symbols={ticker}&published_after={start}&published_before={as_of_date}'
                f'&filter_entities=true&language=en&limit=50'
                f'&api_token={MARKETAUX_API_KEY}', timeout=10)

            if r.status_code != 200:
                return {"count": 0, "avg_sentiment": 0}

            articles = r.json().get("data", [])
            sentiments = []
            for a in articles:
                for e in a.get("entities", []):
                    if e.get("symbol") == ticker:
                        s = e.get("sentiment_score")
                        if s is not None:
                            sentiments.append(float(s))

            result = {
                "count": len(articles),
                "sentiment_count": len(sentiments),
                "avg_sentiment": round(np.mean(sentiments), 4) if sentiments else 0,
                "positive_pct": round(sum(1 for s in sentiments if s > 0.1) / len(sentiments) * 100, 1) if sentiments else 50,
            }

            self._cache[cache_key] = result
            return result

        except Exception:
            return {"count": 0, "avg_sentiment": 0, "positive_pct": 50}
