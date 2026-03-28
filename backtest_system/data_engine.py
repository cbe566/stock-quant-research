"""
數據引擎 — 獲取多市場歷史數據、計算基本面與技術指標
"""

import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import time
warnings.filterwarnings('ignore')


class DataEngine:
    """多市場數據引擎"""

    def __init__(self):
        self.price_cache = {}
        self.info_cache = {}

    # ==================== 價格數據 ====================
    def get_prices(self, ticker, start, end, retry=3):
        """獲取股價數據（含重試機制）"""
        cache_key = f"{ticker}_{start}_{end}"
        if cache_key in self.price_cache:
            return self.price_cache[cache_key]

        for attempt in range(retry):
            try:
                df = yf.download(ticker, start=start, end=end,
                                progress=False, auto_adjust=True)
                if df is not None and len(df) > 0:
                    # 處理多層列名
                    if isinstance(df.columns, pd.MultiIndex):
                        df.columns = df.columns.get_level_values(0)
                    self.price_cache[cache_key] = df
                    return df
            except Exception as e:
                if attempt < retry - 1:
                    time.sleep(2)
                    continue
        return pd.DataFrame()

    def get_stock_info(self, ticker):
        """獲取股票基本面數據"""
        if ticker in self.info_cache:
            return self.info_cache[ticker]

        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            self.info_cache[ticker] = info
            return info
        except:
            return {}

    def get_financials(self, ticker):
        """獲取財務數���"""
        try:
            stock = yf.Ticker(ticker)
            return {
                "income": stock.income_stmt,
                "balance": stock.balance_sheet,
                "cashflow": stock.cashflow,
            }
        except:
            return {}

    # ==================== 技術指標計算 ====================
    @staticmethod
    def calc_sma(series, period):
        return series.rolling(window=period).mean()

    @staticmethod
    def calc_ema(series, period):
        return series.ewm(span=period, adjust=False).mean()

    @staticmethod
    def calc_rsi(series, period=14):
        delta = series.diff()
        gain = delta.where(delta > 0, 0.0)
        loss = -delta.where(delta < 0, 0.0)
        avg_gain = gain.rolling(window=period).mean()
        avg_loss = loss.rolling(window=period).mean()
        rs = avg_gain / avg_loss.replace(0, np.nan)
        rsi = 100 - (100 / (1 + rs))
        return rsi

    @staticmethod
    def calc_macd(series, fast=12, slow=26, signal=9):
        ema_fast = series.ewm(span=fast, adjust=False).mean()
        ema_slow = series.ewm(span=slow, adjust=False).mean()
        macd_line = ema_fast - ema_slow
        signal_line = macd_line.ewm(span=signal, adjust=False).mean()
        histogram = macd_line - signal_line
        return macd_line, signal_line, histogram

    @staticmethod
    def calc_bollinger(series, period=20, std_dev=2):
        sma = series.rolling(window=period).mean()
        std = series.rolling(window=period).std()
        upper = sma + std_dev * std
        lower = sma - std_dev * std
        pct_b = (series - lower) / (upper - lower)
        bandwidth = (upper - lower) / sma
        return upper, sma, lower, pct_b, bandwidth

    @staticmethod
    def calc_atr(high, low, close, period=14):
        tr1 = high - low
        tr2 = abs(high - close.shift(1))
        tr3 = abs(low - close.shift(1))
        tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
        atr = tr.rolling(window=period).mean()
        return atr

    @staticmethod
    def calc_obv(close, volume):
        obv = (np.sign(close.diff()) * volume).fillna(0).cumsum()
        return obv

    @staticmethod
    def calc_vwap(high, low, close, volume):
        typical_price = (high + low + close) / 3
        vwap = (typical_price * volume).cumsum() / volume.cumsum()
        return vwap

    @staticmethod
    def calc_momentum(series, period):
        """計算動量（百分比變化）"""
        return series.pct_change(period)

    @staticmethod
    def calc_volatility(series, period=20):
        """計算歷史波動率（年化）"""
        returns = series.pct_change()
        vol = returns.rolling(window=period).std() * np.sqrt(252)
        return vol

    def calc_all_technicals(self, df):
        """計算所有技術指標"""
        close = df['Close']
        high = df['High']
        low = df['Low']
        volume = df['Volume']

        result = df.copy()

        # 均線
        for p in [5, 10, 20, 50, 100, 200]:
            result[f'SMA_{p}'] = self.calc_sma(close, p)
            result[f'EMA_{p}'] = self.calc_ema(close, p)

        # RSI
        result['RSI_14'] = self.calc_rsi(close, 14)

        # MACD
        result['MACD'], result['MACD_Signal'], result['MACD_Hist'] = self.calc_macd(close)

        # 布林通道
        result['BB_Upper'], result['BB_Mid'], result['BB_Lower'], result['BB_PctB'], result['BB_Width'] = \
            self.calc_bollinger(close)

        # ATR
        result['ATR_14'] = self.calc_atr(high, low, close, 14)

        # OBV
        result['OBV'] = self.calc_obv(close, volume)

        # 動量
        for p in [5, 10, 21, 63, 126, 252]:
            result[f'MOM_{p}'] = self.calc_momentum(close, p)

        # 波動率
        result['VOL_20'] = self.calc_volatility(close, 20)
        result['VOL_60'] = self.calc_volatility(close, 60)

        # 成交量均線
        for p in [5, 20, 60]:
            result[f'VOL_MA_{p}'] = self.calc_sma(volume, p)

        # 量比
        result['VOL_RATIO'] = volume / self.calc_sma(volume, 20)

        # 乖離率 BIAS
        for p in [5, 20, 60]:
            result[f'BIAS_{p}'] = (close - self.calc_sma(close, p)) / self.calc_sma(close, p) * 100

        return result

    # ==================== 基本面指標 ====================
    def get_fundamental_scores(self, ticker):
        """計算基本面評分"""
        info = self.get_stock_info(ticker)
        if not info:
            return {}

        scores = {
            "ticker": ticker,
            "name": info.get("shortName", ""),
            "sector": info.get("sector", ""),
            "industry": info.get("industry", ""),
            "market_cap": info.get("marketCap", 0),

            # 估值指標
            "pe_trailing": info.get("trailingPE", None),
            "pe_forward": info.get("forwardPE", None),
            "pb": info.get("priceToBook", None),
            "ps": info.get("priceToSalesTrailing12Months", None),
            "peg": info.get("pegRatio", None),
            "ev_ebitda": info.get("enterpriseToEbitda", None),

            # 盈利能力
            "roe": info.get("returnOnEquity", None),
            "roa": info.get("returnOnAssets", None),
            "profit_margin": info.get("profitMargins", None),
            "gross_margin": info.get("grossMargins", None),
            "operating_margin": info.get("operatingMargins", None),

            # 成長性
            "revenue_growth": info.get("revenueGrowth", None),
            "earnings_growth": info.get("earningsGrowth", None),
            "earnings_quarterly_growth": info.get("earningsQuarterlyGrowth", None),

            # 財務健康
            "debt_to_equity": info.get("debtToEquity", None),
            "current_ratio": info.get("currentRatio", None),
            "quick_ratio": info.get("quickRatio", None),
            "free_cashflow": info.get("freeCashflow", None),

            # 分紅
            "dividend_yield": info.get("dividendYield", None),
            "payout_ratio": info.get("payoutRatio", None),

            # 分析師
            "target_mean": info.get("targetMeanPrice", None),
            "recommendation": info.get("recommendationKey", None),
            "num_analysts": info.get("numberOfAnalystOpinions", None),
        }

        return scores

    def calc_piotroski_f_score(self, ticker):
        """計算 Piotroski F-Score（0-9分）"""
        info = self.get_stock_info(ticker)
        financials = self.get_financials(ticker)

        score = 0
        details = {}

        try:
            # 1. ROA > 0（盈利能力）
            roa = info.get("returnOnAssets", 0)
            if roa and roa > 0:
                score += 1
                details["roa_positive"] = True

            # 2. 經營現金流 > 0
            fcf = info.get("freeCashflow", 0)
            if fcf and fcf > 0:
                score += 1
                details["cashflow_positive"] = True

            # 3. ROA 改善（用盈餘成長近似）
            eg = info.get("earningsGrowth", 0)
            if eg and eg > 0:
                score += 1
                details["roa_improving"] = True

            # 4. 現金流 > 淨利（盈餘品質）
            # 簡化：用 operating cashflow vs net income
            if fcf and info.get("netIncomeToCommon", 0):
                if fcf > info.get("netIncomeToCommon", 0):
                    score += 1
                    details["accruals_quality"] = True

            # 5. 負債比率下降（用 debt_to_equity 判斷）
            dte = info.get("debtToEquity", None)
            if dte is not None and dte < 100:
                score += 1
                details["leverage_low"] = True

            # 6. 流動比率 > 1
            cr = info.get("currentRatio", 0)
            if cr and cr > 1:
                score += 1
                details["liquidity_good"] = True

            # 7. 未稀釋（簡化：檢查股數是否增加 — 這裡用股份數判斷）
            shares = info.get("sharesOutstanding", 0)
            if shares and shares > 0:
                score += 1  # 簡化處理
                details["no_dilution"] = True

            # 8. 毛利率改善
            gm = info.get("grossMargins", 0)
            if gm and gm > 0.3:
                score += 1
                details["gross_margin_good"] = True

            # 9. 資產周轉率（用營收成長近似）
            rg = info.get("revenueGrowth", 0)
            if rg and rg > 0:
                score += 1
                details["asset_turnover_improving"] = True

        except Exception:
            pass

        return {"score": score, "max_score": 9, "details": details}
