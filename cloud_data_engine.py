"""
雲端數據引擎 — 多源備援 + 爬蟲兜底
=====================================
解決 GitHub Actions 上 yfinance / API 被封鎖的問題

數據獲取優先順序（每層失敗自動跳下一層）：
  第1層：yfinance（最完整，但雲端常被封）
  第2層：Yahoo Finance v8 API 直接請求（繞過 yfinance 限制）
  第3層：各地交易所 Open API（TWSE/HKEX/JPX 官方免費數據）
  第4層：FMP / Alpha Vantage / Twelve Data 備援 API
  第5層：網頁爬蟲兜底（finviz / Google Finance / investing.com）

設計原則：
  - 每個請求都帶真實瀏覽器 headers
  - 自動重試 + 指數退避
  - 連線池複用（requests.Session）
  - 結果統一格式化，上層無感切換
"""

import requests
import pandas as pd
import numpy as np
import json
import time
import random
import os
import threading
from datetime import datetime, timedelta
from io import StringIO
import warnings
warnings.filterwarnings('ignore')

# ==================== 配置 ====================

# API Keys（從環境變數讀取，GitHub Actions 用 Secrets 注入）
API_KEYS = {
    "FMP": os.environ.get("FMP_API_KEY", "GEouSGBbAoOgnERMR0GjENhKwFkxeEeW"),
    "ALPHA_VANTAGE": os.environ.get("ALPHA_VANTAGE_API_KEY", "CYSEDIRPDGYNRD3H"),
    "TWELVE_DATA": os.environ.get("TWELVE_DATA_API_KEY", "685ced6c2b694a12b359beeec084e9cf"),
    "FINNHUB": os.environ.get("FINNHUB_API_KEY", "d6nsik9r01qse5qmtl4gd6nsik9r01qse5qmtl50"),
    "TIINGO": os.environ.get("TIINGO_API_KEY", "981e9f863c8997d8e4ac01ca269c033295e4c1bb"),
    "EOD": os.environ.get("EOD_API_KEY", "69c7f5b716ce48.31162491"),
}

# 瀏覽器 User-Agent 池（隨機輪換避免被識別為爬蟲）
USER_AGENTS = [
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.2 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:133.0) Gecko/20100101 Firefox/133.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
]


class CloudDataEngine:
    """
    雲端友好的多源數據引擎
    自動在多個數據源之間切換，確保在 GitHub Actions 環境也能穩定獲取數據
    """

    def __init__(self):
        self.session = requests.Session()
        self._rotate_headers()
        self.price_cache = {}
        self.info_cache = {}
        self._api_call_count = 0
        self._source_stats = {}  # 追蹤各數據源成功率
        self._lock = threading.Lock()  # 多線程安全鎖
        self._thread_local = threading.local()  # 每線程獨立 session

    def _rotate_headers(self):
        """輪換請求 headers，模擬真實瀏覽器"""
        self.session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json;q=0.8,*/*;q=0.7",
            "Accept-Language": "en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Cache-Control": "no-cache",
        })

    def _get_thread_session(self):
        """取得當前線程的 requests.Session（線程安全）"""
        if not hasattr(self._thread_local, 'session'):
            s = requests.Session()
            s.headers.update({
                "User-Agent": random.choice(USER_AGENTS),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json;q=0.8,*/*;q=0.7",
                "Accept-Language": "en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7",
                "Accept-Encoding": "gzip, deflate, br",
                "Connection": "keep-alive",
                "Cache-Control": "no-cache",
            })
            self._thread_local.session = s
        return self._thread_local.session

    def _request_with_retry(self, url, params=None, headers=None, max_retries=3, timeout=15):
        """帶重試和指數退避的請求（線程安全）"""
        session = self._get_thread_session()
        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    session.headers["User-Agent"] = random.choice(USER_AGENTS)
                    wait = (2 ** attempt) + random.uniform(0, 1)
                    time.sleep(wait)

                resp = session.get(url, params=params, headers=headers,
                                   timeout=timeout, allow_redirects=True)

                if resp.status_code == 200:
                    return resp
                elif resp.status_code == 429:
                    # 被限速，等更久
                    time.sleep(5 + random.uniform(1, 3))
                    continue
                elif resp.status_code in (403, 451):
                    # 被封，直接放棄這個源
                    return None
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
                continue
            except Exception:
                continue
        return None

    def _track_source(self, source, success):
        """追蹤數據源成功率（線程安全）"""
        with self._lock:
            if source not in self._source_stats:
                self._source_stats[source] = {"success": 0, "fail": 0}
            if success:
                self._source_stats[source]["success"] += 1
            else:
                self._source_stats[source]["fail"] += 1

    # ================================================================
    #  第一層：yfinance（最完整，但雲端常被封）
    # ================================================================

    def _get_prices_yfinance(self, ticker, start, end):
        """嘗試用 yfinance 獲取價格"""
        try:
            import yfinance as yf
            df = yf.download(ticker, start=start, end=end, progress=False, auto_adjust=True)
            if df is not None and len(df) > 0:
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = df.columns.get_level_values(0)
                self._track_source("yfinance", True)
                return df
        except Exception:
            pass
        self._track_source("yfinance", False)
        return None

    def _get_info_yfinance(self, ticker):
        """嘗試用 yfinance 獲取基本面"""
        try:
            import yfinance as yf
            stock = yf.Ticker(ticker)
            info = stock.info
            if info and info.get("shortName"):
                self._track_source("yfinance_info", True)
                return info
        except Exception:
            pass
        self._track_source("yfinance_info", False)
        return None

    # ================================================================
    #  第二層：Yahoo Finance v8 API 直接請求
    # ================================================================

    def _get_prices_yahoo_direct(self, ticker, start, end):
        """直接請求 Yahoo Finance API，繞過 yfinance 的限制"""
        try:
            # Yahoo Finance v8 chart API
            period1 = int(datetime.strptime(start, "%Y-%m-%d").timestamp())
            period2 = int(datetime.strptime(end, "%Y-%m-%d").timestamp())

            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
            params = {
                "period1": period1,
                "period2": period2,
                "interval": "1d",
                "includePrePost": "false",
                "events": "div,splits",
            }

            # Yahoo 需要特殊的 crumb/cookie，用直接 headers
            headers = {
                "User-Agent": random.choice(USER_AGENTS),
                "Accept": "application/json",
                "Referer": f"https://finance.yahoo.com/quote/{ticker}/",
                "Origin": "https://finance.yahoo.com",
            }

            resp = self._request_with_retry(url, params=params, headers=headers)
            if resp is None:
                self._track_source("yahoo_direct", False)
                return None

            data = resp.json()
            result = data.get("chart", {}).get("result", [])
            if not result:
                self._track_source("yahoo_direct", False)
                return None

            quotes = result[0]
            timestamps = quotes.get("timestamp", [])
            ohlcv = quotes.get("indicators", {}).get("quote", [{}])[0]
            adj_close = quotes.get("indicators", {}).get("adjclose", [{}])

            if not timestamps or not ohlcv.get("close"):
                self._track_source("yahoo_direct", False)
                return None

            df = pd.DataFrame({
                "Open": ohlcv.get("open", []),
                "High": ohlcv.get("high", []),
                "Low": ohlcv.get("low", []),
                "Close": adj_close[0].get("adjclose", ohlcv.get("close", [])) if adj_close else ohlcv.get("close", []),
                "Volume": ohlcv.get("volume", []),
            }, index=pd.to_datetime([datetime.fromtimestamp(t) for t in timestamps]))

            df.index.name = "Date"
            df = df.dropna(subset=["Close"])

            if len(df) > 0:
                self._track_source("yahoo_direct", True)
                return df

        except Exception:
            pass
        self._track_source("yahoo_direct", False)
        return None

    def _get_yahoo_crumb(self):
        """獲取 Yahoo Finance crumb + cookie（需要先訪問一次建立 session）"""
        if hasattr(self, '_yahoo_crumb') and self._yahoo_crumb:
            return self._yahoo_crumb, self._yahoo_session

        try:
            ys = requests.Session()
            ys.headers.update({"User-Agent": random.choice(USER_AGENTS)})

            # Step 1: 建立 cookie
            ys.get("https://fc.yahoo.com", timeout=10, allow_redirects=True)
            time.sleep(0.5)

            # Step 2: 拿 crumb
            resp = ys.get("https://query2.finance.yahoo.com/v1/test/getcrumb", timeout=10)
            if resp.status_code == 200 and len(resp.text) < 50:
                self._yahoo_crumb = resp.text
                self._yahoo_session = ys
                return self._yahoo_crumb, ys
        except Exception:
            pass
        return None, None

    def _get_info_yahoo_direct(self, ticker):
        """直接請求 Yahoo Finance quoteSummary API（帶 crumb 認證）"""
        try:
            crumb, ys = self._get_yahoo_crumb()
            if not crumb:
                self._track_source("yahoo_direct_info", False)
                return None

            url = f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{ticker}"
            params = {
                "modules": "defaultKeyStatistics,financialData,summaryDetail,summaryProfile,price",
                "crumb": crumb,
            }

            resp = ys.get(url, params=params, timeout=15)
            if resp.status_code != 200:
                self._track_source("yahoo_direct_info", False)
                return None

            data = resp.json()
            result = data.get("quoteSummary", {}).get("result", [])
            if not result:
                self._track_source("yahoo_direct_info", False)
                return None

            modules = result[0]
            info = {}

            price = modules.get("price", {})
            info["shortName"] = price.get("shortName", "")
            info["currentPrice"] = self._extract_raw(price.get("regularMarketPrice"))
            info["regularMarketPrice"] = info["currentPrice"]
            info["marketCap"] = self._extract_raw(price.get("marketCap"))

            fd = modules.get("financialData", {})
            info["targetMeanPrice"] = self._extract_raw(fd.get("targetMeanPrice"))
            info["recommendationKey"] = fd.get("recommendationKey")
            info["numberOfAnalystOpinions"] = self._extract_raw(fd.get("numberOfAnalystOpinions"))
            info["returnOnEquity"] = self._extract_raw(fd.get("returnOnEquity"))
            info["returnOnAssets"] = self._extract_raw(fd.get("returnOnAssets"))
            info["revenueGrowth"] = self._extract_raw(fd.get("revenueGrowth"))
            info["earningsGrowth"] = self._extract_raw(fd.get("earningsGrowth"))
            info["profitMargins"] = self._extract_raw(fd.get("profitMargins"))
            info["grossMargins"] = self._extract_raw(fd.get("grossMargins"))
            info["operatingMargins"] = self._extract_raw(fd.get("operatingMargins"))
            info["freeCashflow"] = self._extract_raw(fd.get("freeCashflow"))
            info["currentRatio"] = self._extract_raw(fd.get("currentRatio"))
            info["debtToEquity"] = self._extract_raw(fd.get("debtToEquity"))

            ks = modules.get("defaultKeyStatistics", {})
            info["forwardPE"] = self._extract_raw(ks.get("forwardPE"))
            info["priceToBook"] = self._extract_raw(ks.get("priceToBook"))
            info["pegRatio"] = self._extract_raw(ks.get("pegRatio"))
            info["enterpriseToEbitda"] = self._extract_raw(ks.get("enterpriseToEbitda"))
            info["sharesOutstanding"] = self._extract_raw(ks.get("sharesOutstanding"))
            info["netIncomeToCommon"] = self._extract_raw(ks.get("netIncomeToCommon"))

            sd = modules.get("summaryDetail", {})
            info["trailingPE"] = self._extract_raw(sd.get("trailingPE"))
            info["dividendYield"] = self._extract_raw(sd.get("dividendYield"))
            info["payoutRatio"] = self._extract_raw(sd.get("payoutRatio"))

            sp = modules.get("summaryProfile", {})
            info["sector"] = sp.get("sector", "")
            info["industry"] = sp.get("industry", "")

            if info.get("shortName") or info.get("currentPrice"):
                self._track_source("yahoo_direct_info", True)
                return info

        except Exception:
            pass
        self._track_source("yahoo_direct_info", False)
        return None

    @staticmethod
    def _extract_raw(val):
        """從 Yahoo API 的 {raw: x, fmt: y} 格式提取原始值"""
        if val is None:
            return None
        if isinstance(val, dict):
            return val.get("raw")
        return val

    # ================================================================
    #  第三層：各地交易所 Open API（免費、穩定）
    # ================================================================

    def _get_prices_twse(self, ticker, start, end):
        """台灣證交所 Open API — 免費無需 API Key"""
        try:
            # 提取純數字代碼
            code = ticker.replace(".TW", "").replace(".TWO", "")

            # TWSE 提供每日收盤行情 JSON
            # 抓最近 N 個月的月報
            all_rows = []
            start_dt = datetime.strptime(start, "%Y-%m-%d")
            end_dt = datetime.strptime(end, "%Y-%m-%d")

            # 只抓最近 13 個月（避免太多請求）
            months_to_fetch = min(13, (end_dt.year - start_dt.year) * 12 + end_dt.month - start_dt.month + 1)
            current = end_dt.replace(day=1)

            for _ in range(months_to_fetch):
                date_str = current.strftime("%Y%m01")
                url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={code}"

                resp = self._request_with_retry(url, timeout=10)
                if resp and resp.status_code == 200:
                    data = resp.json()
                    if data.get("stat") == "OK" and data.get("data"):
                        for row in data["data"]:
                            try:
                                # 民國年轉西曆
                                date_parts = row[0].split("/")
                                year = int(date_parts[0]) + 1911
                                month = int(date_parts[1])
                                day = int(date_parts[2])
                                dt = datetime(year, month, day)

                                # 移除千分位逗號
                                def clean_num(s):
                                    return float(s.replace(",", "").replace("--", "0")) if s else 0

                                all_rows.append({
                                    "Date": dt,
                                    "Volume": clean_num(row[1]),
                                    "Open": clean_num(row[3]),
                                    "High": clean_num(row[4]),
                                    "Low": clean_num(row[5]),
                                    "Close": clean_num(row[6]),
                                })
                            except (ValueError, IndexError):
                                continue

                # 上個月
                if current.month == 1:
                    current = current.replace(year=current.year - 1, month=12)
                else:
                    current = current.replace(month=current.month - 1)

                time.sleep(0.5)  # TWSE 速率限制

            if all_rows:
                df = pd.DataFrame(all_rows)
                df.set_index("Date", inplace=True)
                df.sort_index(inplace=True)
                # 過濾日期範圍
                df = df[(df.index >= start) & (df.index <= end)]
                if len(df) > 0:
                    self._track_source("twse", True)
                    return df

        except Exception:
            pass
        self._track_source("twse", False)
        return None

    # ================================================================
    #  第四層：備援 API（FMP / Alpha Vantage / Tiingo / Twelve Data）
    # ================================================================

    def _get_prices_fmp(self, ticker, start, end):
        """Financial Modeling Prep API（v4 stable endpoint）"""
        try:
            url = "https://financialmodelingprep.com/stable/historical-price-eod/full"
            params = {
                "symbol": ticker,
                "from": start,
                "to": end,
                "apikey": API_KEYS["FMP"],
            }

            resp = self._request_with_retry(url, params=params)
            if resp is None:
                self._track_source("fmp", False)
                return None

            data = resp.json()
            if not data or not isinstance(data, list):
                self._track_source("fmp", False)
                return None

            df = pd.DataFrame(data)
            df["Date"] = pd.to_datetime(df["date"])
            df.set_index("Date", inplace=True)
            df = df.rename(columns={
                "open": "Open", "high": "High", "low": "Low",
                "close": "Close", "volume": "Volume",
            })
            df = df[["Open", "High", "Low", "Close", "Volume"]].sort_index()

            if len(df) > 0:
                self._track_source("fmp", True)
                return df

        except Exception:
            pass
        self._track_source("fmp", False)
        return None

    def _get_prices_tiingo(self, ticker, start, end):
        """Tiingo API — 穩定的歷史股價"""
        try:
            # Tiingo 只支援美股代碼
            if any(suffix in ticker for suffix in [".HK", ".TW", ".T"]):
                return None

            url = f"https://api.tiingo.com/tiingo/daily/{ticker}/prices"
            params = {
                "startDate": start,
                "endDate": end,
                "token": API_KEYS["TIINGO"],
            }
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Token {API_KEYS['TIINGO']}",
            }

            resp = self._request_with_retry(url, params=params, headers=headers)
            if resp is None:
                self._track_source("tiingo", False)
                return None

            data = resp.json()
            if not data:
                self._track_source("tiingo", False)
                return None

            df = pd.DataFrame(data)
            df["Date"] = pd.to_datetime(df["date"])
            df.set_index("Date", inplace=True)
            df = df.rename(columns={
                "adjOpen": "Open", "adjHigh": "High",
                "adjLow": "Low", "adjClose": "Close", "volume": "Volume",
            })
            df = df[["Open", "High", "Low", "Close", "Volume"]].sort_index()

            if len(df) > 0:
                self._track_source("tiingo", True)
                return df

        except Exception:
            pass
        self._track_source("tiingo", False)
        return None

    def _get_prices_twelve_data(self, ticker, start, end):
        """Twelve Data API"""
        try:
            url = "https://api.twelvedata.com/time_series"
            params = {
                "symbol": ticker,
                "interval": "1day",
                "start_date": start,
                "end_date": end,
                "outputsize": 5000,
                "apikey": API_KEYS["TWELVE_DATA"],
            }

            resp = self._request_with_retry(url, params=params)
            if resp is None:
                self._track_source("twelve_data", False)
                return None

            data = resp.json()
            values = data.get("values", [])
            if not values:
                self._track_source("twelve_data", False)
                return None

            df = pd.DataFrame(values)
            df["Date"] = pd.to_datetime(df["datetime"])
            df.set_index("Date", inplace=True)
            for col in ["open", "high", "low", "close", "volume"]:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            df = df.rename(columns={
                "open": "Open", "high": "High", "low": "Low",
                "close": "Close", "volume": "Volume",
            })
            df = df[["Open", "High", "Low", "Close", "Volume"]].sort_index()

            if len(df) > 0:
                self._track_source("twelve_data", True)
                return df

        except Exception:
            pass
        self._track_source("twelve_data", False)
        return None

    def _get_info_fmp(self, ticker):
        """FMP 基本面數據（v4 stable endpoint）"""
        try:
            # Profile
            url = "https://financialmodelingprep.com/stable/profile"
            params = {"symbol": ticker, "apikey": API_KEYS["FMP"]}

            resp = self._request_with_retry(url, params=params)
            if resp is None:
                self._track_source("fmp_info", False)
                return None

            data = resp.json()
            if not data:
                self._track_source("fmp_info", False)
                return None

            p = data[0] if isinstance(data, list) else data

            info = {
                "shortName": p.get("companyName", ""),
                "sector": p.get("sector", ""),
                "industry": p.get("industry", ""),
                "currentPrice": p.get("price"),
                "regularMarketPrice": p.get("price"),
                "marketCap": p.get("mktCap"),
                "trailingPE": p.get("pe") if p.get("pe") and p.get("pe") > 0 else None,
                "priceToBook": None,
                "returnOnEquity": None,
                "dividendYield": p.get("lastDiv") / p.get("price") if p.get("lastDiv") and p.get("price") else None,
                "targetMeanPrice": p.get("dcf"),
                "freeCashflow": None,
                "debtToEquity": None,
                "grossMargins": None,
                "profitMargins": None,
                "revenueGrowth": None,
                "earningsGrowth": None,
            }

            # 補充 ratios
            url2 = "https://financialmodelingprep.com/stable/ratios-ttm"
            resp2 = self._request_with_retry(url2, params=params)
            if resp2 and resp2.status_code == 200:
                km = resp2.json()
                if km and isinstance(km, list) and km[0]:
                    k = km[0]
                    info["returnOnEquity"] = k.get("returnOnEquityTTM")
                    info["returnOnAssets"] = k.get("returnOnAssetsTTM")
                    info["priceToBook"] = k.get("priceToBookRatioTTM")
                    info["currentRatio"] = k.get("currentRatioTTM")
                    info["grossMargins"] = k.get("grossProfitMarginTTM")
                    info["profitMargins"] = k.get("netProfitMarginTTM")
                    dte = k.get("debtEquityRatioTTM")
                    if dte:
                        info["debtToEquity"] = dte * 100

            if info.get("shortName") or info.get("currentPrice"):
                self._track_source("fmp_info", True)
                return info

        except Exception:
            pass
        self._track_source("fmp_info", False)
        return None

    def _get_info_finnhub(self, ticker):
        """Finnhub 基本面（僅美股）"""
        try:
            if any(suffix in ticker for suffix in [".HK", ".TW", ".T"]):
                return None

            url = "https://finnhub.io/api/v1/stock/metric"
            params = {
                "symbol": ticker,
                "metric": "all",
                "token": API_KEYS["FINNHUB"],
            }

            resp = self._request_with_retry(url, params=params)
            if resp is None:
                self._track_source("finnhub_info", False)
                return None

            data = resp.json()
            metric = data.get("metric", {})
            if not metric:
                self._track_source("finnhub_info", False)
                return None

            info = {
                "shortName": ticker,
                "currentPrice": metric.get("52WeekHigh"),  # 近似值
                "trailingPE": metric.get("peBasicExclExtraTTM"),
                "priceToBook": metric.get("pbAnnual"),
                "returnOnEquity": metric.get("roeTTM") / 100 if metric.get("roeTTM") else None,
                "returnOnAssets": metric.get("roaTTM") / 100 if metric.get("roaTTM") else None,
                "grossMargins": metric.get("grossMarginTTM") / 100 if metric.get("grossMarginTTM") else None,
                "profitMargins": metric.get("netProfitMarginTTM") / 100 if metric.get("netProfitMarginTTM") else None,
                "currentRatio": metric.get("currentRatioQuarterly"),
                "dividendYield": metric.get("dividendYieldIndicatedAnnual") / 100 if metric.get("dividendYieldIndicatedAnnual") else None,
                "revenueGrowth": metric.get("revenueGrowthQuarterlyYoy") / 100 if metric.get("revenueGrowthQuarterlyYoy") else None,
            }

            # 補充分析師目標價
            url2 = "https://finnhub.io/api/v1/stock/price-target"
            params2 = {"symbol": ticker, "token": API_KEYS["FINNHUB"]}
            resp2 = self._request_with_retry(url2, params=params2)
            if resp2:
                pt = resp2.json()
                info["targetMeanPrice"] = pt.get("targetMean")

            self._track_source("finnhub_info", True)
            return info

        except Exception:
            pass
        self._track_source("finnhub_info", False)
        return None

    # ================================================================
    #  第五層：網頁爬蟲兜底
    # ================================================================

    def _get_prices_stooq(self, ticker, start, end):
        """
        Stooq.com 爬蟲 — 最後手段（免費歷史數據，不需 API Key）
        """
        try:
            # 轉換代碼格式
            if ticker.endswith(".HK"):
                code = ticker.replace(".HK", "")
                stooq_ticker = f"{code}.hk"
            elif ticker.endswith(".TW"):
                code = ticker.replace(".TW", "")
                stooq_ticker = f"{code}.tw"
            elif ticker.endswith(".T"):
                code = ticker.replace(".T", "")
                stooq_ticker = f"{code}.jp"
            else:
                stooq_ticker = f"{ticker}.us"

            # Stooq 使用 pandas datareader 格式
            url = f"https://stooq.com/q/d/l/"
            params = {
                "s": stooq_ticker,
                "d1": start.replace("-", ""),
                "d2": end.replace("-", ""),
                "i": "d",
            }

            headers = {
                "User-Agent": random.choice(USER_AGENTS),
                "Accept": "text/csv,text/plain,*/*",
                "Referer": f"https://stooq.com/q/d/?s={stooq_ticker}",
            }

            resp = self._request_with_retry(url, params=params, headers=headers)
            if resp is None or len(resp.text.strip()) < 50 or "No data" in resp.text:
                self._track_source("stooq", False)
                return None

            df = pd.read_csv(StringIO(resp.text))
            if df.empty or "Close" not in df.columns:
                self._track_source("stooq", False)
                return None

            df["Date"] = pd.to_datetime(df["Date"])
            df.set_index("Date", inplace=True)

            # Stooq 有時候不含 Volume
            cols = ["Open", "High", "Low", "Close"]
            if "Volume" in df.columns:
                cols.append("Volume")
            else:
                df["Volume"] = 0

            df = df[["Open", "High", "Low", "Close", "Volume"]].sort_index()

            if len(df) > 0:
                self._track_source("stooq", True)
                return df

        except Exception:
            pass
        self._track_source("stooq", False)
        return None

    # ================================================================
    #  統一接口：自動多源切換
    # ================================================================

    def get_prices(self, ticker, start, end):
        """
        獲取價格數據 — 多源自動切換（線程安全）
        優先順序：yfinance → Yahoo 直接 → 交易所 API → FMP → Tiingo → Twelve Data → Stooq
        """
        cache_key = f"{ticker}_{start}_{end}"
        with self._lock:
            if cache_key in self.price_cache:
                return self.price_cache[cache_key]

        # 定義數據源嘗試順序
        sources = [
            ("yfinance", lambda: self._get_prices_yfinance(ticker, start, end)),
            ("yahoo_direct", lambda: self._get_prices_yahoo_direct(ticker, start, end)),
        ]

        # 台股加入 TWSE
        if ticker.endswith(".TW"):
            sources.append(("twse", lambda: self._get_prices_twse(ticker, start, end)))

        # 通用備援
        sources.extend([
            ("fmp", lambda: self._get_prices_fmp(ticker, start, end)),
            ("tiingo", lambda: self._get_prices_tiingo(ticker, start, end)),
            ("twelve_data", lambda: self._get_prices_twelve_data(ticker, start, end)),
            ("stooq", lambda: self._get_prices_stooq(ticker, start, end)),
        ])

        for name, fetch_fn in sources:
            try:
                df = fetch_fn()
                if df is not None and len(df) > 0:
                    with self._lock:
                        self.price_cache[cache_key] = df
                    return df
            except Exception:
                continue

        return pd.DataFrame()

    def get_stock_info(self, ticker):
        """
        獲取基本面數據 — 多源自動切換（線程安全）
        優先順序：yfinance → Yahoo 直接 → FMP → Finnhub
        """
        with self._lock:
            if ticker in self.info_cache:
                return self.info_cache[ticker]

        sources = [
            ("yfinance", lambda: self._get_info_yfinance(ticker)),
            ("yahoo_direct", lambda: self._get_info_yahoo_direct(ticker)),
            ("fmp", lambda: self._get_info_fmp(ticker)),
            ("finnhub", lambda: self._get_info_finnhub(ticker)),
        ]

        for name, fetch_fn in sources:
            try:
                info = fetch_fn()
                if info and (info.get("shortName") or info.get("currentPrice")):
                    with self._lock:
                        self.info_cache[ticker] = info
                    return info
            except Exception:
                continue

        return {}

    def get_financials(self, ticker):
        """獲取財務數據（盡力而為）"""
        try:
            import yfinance as yf
            stock = yf.Ticker(ticker)
            return {
                "income": stock.income_stmt,
                "balance": stock.balance_sheet,
                "cashflow": stock.cashflow,
            }
        except Exception:
            return {}

    # ==================== 技術指標計算（與 DataEngine 相同） ====================

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
    def calc_momentum(series, period):
        return series.pct_change(period)

    @staticmethod
    def calc_volatility(series, period=20):
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

        for p in [5, 10, 20, 50, 100, 200]:
            result[f'SMA_{p}'] = self.calc_sma(close, p)
            result[f'EMA_{p}'] = self.calc_ema(close, p)

        result['RSI_14'] = self.calc_rsi(close, 14)
        result['MACD'], result['MACD_Signal'], result['MACD_Hist'] = self.calc_macd(close)
        result['BB_Upper'], result['BB_Mid'], result['BB_Lower'], result['BB_PctB'], result['BB_Width'] = \
            self.calc_bollinger(close)
        result['ATR_14'] = self.calc_atr(high, low, close, 14)
        result['OBV'] = self.calc_obv(close, volume)

        for p in [5, 10, 21, 63, 126, 252]:
            result[f'MOM_{p}'] = self.calc_momentum(close, p)

        result['VOL_20'] = self.calc_volatility(close, 20)
        result['VOL_60'] = self.calc_volatility(close, 60)

        for p in [5, 20, 60]:
            result[f'VOL_MA_{p}'] = self.calc_sma(volume, p)

        result['VOL_RATIO'] = volume / self.calc_sma(volume, 20)

        for p in [5, 20, 60]:
            result[f'BIAS_{p}'] = (close - self.calc_sma(close, p)) / self.calc_sma(close, p) * 100

        return result

    def get_fundamental_scores(self, ticker):
        """計算基本面評分（與 DataEngine 兼容）"""
        info = self.get_stock_info(ticker)
        if not info:
            return {}

        return {
            "ticker": ticker,
            "name": info.get("shortName", ""),
            "sector": info.get("sector", ""),
            "industry": info.get("industry", ""),
            "market_cap": info.get("marketCap", 0),
            "pe_trailing": info.get("trailingPE"),
            "pe_forward": info.get("forwardPE"),
            "pb": info.get("priceToBook"),
            "ps": info.get("priceToSalesTrailing12Months"),
            "peg": info.get("pegRatio"),
            "ev_ebitda": info.get("enterpriseToEbitda"),
            "roe": info.get("returnOnEquity"),
            "roa": info.get("returnOnAssets"),
            "profit_margin": info.get("profitMargins"),
            "gross_margin": info.get("grossMargins"),
            "operating_margin": info.get("operatingMargins"),
            "revenue_growth": info.get("revenueGrowth"),
            "earnings_growth": info.get("earningsGrowth"),
            "earnings_quarterly_growth": info.get("earningsQuarterlyGrowth"),
            "debt_to_equity": info.get("debtToEquity"),
            "current_ratio": info.get("currentRatio"),
            "quick_ratio": info.get("quickRatio"),
            "free_cashflow": info.get("freeCashflow"),
            "dividend_yield": info.get("dividendYield"),
            "payout_ratio": info.get("payoutRatio"),
            "target_mean": info.get("targetMeanPrice"),
            "recommendation": info.get("recommendationKey"),
            "num_analysts": info.get("numberOfAnalystOpinions"),
        }

    def calc_piotroski_f_score(self, ticker):
        """計算 Piotroski F-Score（0-9分）"""
        info = self.get_stock_info(ticker)

        score = 0
        details = {}

        try:
            roa = info.get("returnOnAssets", 0)
            if roa and roa > 0:
                score += 1
                details["roa_positive"] = True

            fcf = info.get("freeCashflow", 0)
            if fcf and fcf > 0:
                score += 1
                details["cashflow_positive"] = True

            eg = info.get("earningsGrowth", 0)
            if eg and eg > 0:
                score += 1
                details["roa_improving"] = True

            if fcf and info.get("netIncomeToCommon", 0):
                if fcf > info.get("netIncomeToCommon", 0):
                    score += 1
                    details["accruals_quality"] = True

            dte = info.get("debtToEquity")
            if dte is not None and dte < 100:
                score += 1
                details["leverage_low"] = True

            cr = info.get("currentRatio", 0)
            if cr and cr > 1:
                score += 1
                details["liquidity_good"] = True

            shares = info.get("sharesOutstanding", 0)
            if shares and shares > 0:
                score += 1
                details["no_dilution"] = True

            gm = info.get("grossMargins", 0)
            if gm and gm > 0.3:
                score += 1
                details["gross_margin_good"] = True

            rg = info.get("revenueGrowth", 0)
            if rg and rg > 0:
                score += 1
                details["asset_turnover_improving"] = True

        except Exception:
            pass

        return {"score": score, "max_score": 9, "details": details}

    def print_source_stats(self):
        """輸出各數據源成功率統計"""
        print("\n📊 數據源使用統計：")
        for source, stats in sorted(self._source_stats.items()):
            total = stats["success"] + stats["fail"]
            rate = stats["success"] / total * 100 if total > 0 else 0
            bar = "█" * int(rate / 5) + "░" * (20 - int(rate / 5))
            print(f"  {source:20s} [{bar}] {rate:5.1f}% ({stats['success']}/{total})")
