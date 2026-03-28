"""
Finnhub 數據引擎 — 提供分析師評級、內部人交易、財務指標、新聞情緒
"""

import requests
import time
from datetime import datetime, timedelta
from config import FINNHUB_API_KEY

BASE_URL = "https://finnhub.io/api/v1"


class FinnhubEngine:
    """Finnhub 數據引擎"""

    def __init__(self, api_key=FINNHUB_API_KEY):
        self.api_key = api_key
        self.cache = {}
        self._call_count = 0

    def _get(self, endpoint, params=None):
        """帶速率控制的 API 請求"""
        if params is None:
            params = {}
        params["token"] = self.api_key

        # 速率控制：60次/分鐘
        self._call_count += 1
        if self._call_count % 55 == 0:
            time.sleep(2)

        try:
            r = requests.get(f"{BASE_URL}/{endpoint}", params=params, timeout=10)
            if r.status_code == 200:
                return r.json()
            elif r.status_code == 429:
                time.sleep(5)
                return self._get(endpoint, params)
        except Exception:
            pass
        return None

    # ==================== 分析師評級 ====================
    def get_analyst_ratings(self, ticker):
        """獲取分析師評級（買入/持有/賣出數量）"""
        cache_key = f"rating_{ticker}"
        if cache_key in self.cache:
            return self.cache[cache_key]

        data = self._get("stock/recommendation", {"symbol": ticker})
        if not data or not isinstance(data, list) or len(data) == 0:
            return None

        latest = data[0]
        total = (latest.get("strongBuy", 0) + latest.get("buy", 0) +
                 latest.get("hold", 0) + latest.get("sell", 0) +
                 latest.get("strongSell", 0))

        result = {
            "strong_buy": latest.get("strongBuy", 0),
            "buy": latest.get("buy", 0),
            "hold": latest.get("hold", 0),
            "sell": latest.get("sell", 0),
            "strong_sell": latest.get("strongSell", 0),
            "total_analysts": total,
            "bull_pct": round((latest.get("strongBuy", 0) + latest.get("buy", 0)) / total * 100, 1) if total > 0 else 0,
            "period": latest.get("period", ""),
            "consensus": self._calc_consensus(latest),
        }

        self.cache[cache_key] = result
        return result

    @staticmethod
    def _calc_consensus(rating):
        sb = rating.get("strongBuy", 0)
        b = rating.get("buy", 0)
        h = rating.get("hold", 0)
        s = rating.get("sell", 0)
        ss = rating.get("strongSell", 0)
        total = sb + b + h + s + ss
        if total == 0:
            return "無數據"
        # 加權平均：5=強烈買入, 4=買入, 3=持有, 2=賣出, 1=強烈賣出
        score = (sb * 5 + b * 4 + h * 3 + s * 2 + ss * 1) / total
        if score >= 4.2: return "強烈買入"
        elif score >= 3.5: return "買入"
        elif score >= 2.5: return "持有"
        elif score >= 1.8: return "賣出"
        else: return "強烈賣出"

    def get_price_target(self, ticker):
        """獲取分析師目標價"""
        data = self._get("stock/price-target", {"symbol": ticker})
        if not data:
            return None
        return {
            "target_high": data.get("targetHigh"),
            "target_low": data.get("targetLow"),
            "target_mean": data.get("targetMean"),
            "target_median": data.get("targetMedian"),
            "last_updated": data.get("lastUpdated"),
        }

    # ==================== 內部人交易 ====================
    def get_insider_transactions(self, ticker, months=6):
        """獲取內部人交易（近N個月）"""
        cache_key = f"insider_{ticker}_{months}"
        if cache_key in self.cache:
            return self.cache[cache_key]

        from_date = (datetime.now() - timedelta(days=months * 30)).strftime("%Y-%m-%d")
        data = self._get("stock/insider-transactions", {"symbol": ticker, "from": from_date})

        if not data or "data" not in data:
            return None

        transactions = data["data"]
        if not transactions:
            return {"total": 0, "net_signal": "無數據"}

        buys = sum(1 for t in transactions
                   if t.get("transactionType") in ["P - Purchase", "A - Grant"])
        sells = sum(1 for t in transactions
                    if t.get("transactionType") in ["S - Sale", "S - Sale+OE"])

        buy_value = sum(abs(t.get("share", 0) * t.get("price", 0))
                       for t in transactions
                       if t.get("transactionType") in ["P - Purchase"])
        sell_value = sum(abs(t.get("share", 0) * t.get("price", 0))
                        for t in transactions
                        if t.get("transactionType") in ["S - Sale", "S - Sale+OE"])

        # 判斷信號
        if buys > sells * 2:
            signal = "強烈看多（內部人大量買入）"
        elif buys > sells:
            signal = "偏多（內部人淨買入）"
        elif sells > buys * 3:
            signal = "看空（內部人大量賣出）"
        elif sells > buys:
            signal = "偏空（內部人淨賣出）"
        else:
            signal = "中性"

        result = {
            "total": len(transactions),
            "buys": buys,
            "sells": sells,
            "buy_value": round(buy_value, 0),
            "sell_value": round(sell_value, 0),
            "net_signal": signal,
            "recent_3": transactions[:3],  # 最近3筆
        }

        self.cache[cache_key] = result
        return result

    # ==================== 財務指標（132個）====================
    def get_financial_metrics(self, ticker):
        """獲取完整財務指標"""
        cache_key = f"metrics_{ticker}"
        if cache_key in self.cache:
            return self.cache[cache_key]

        data = self._get("stock/metric", {"symbol": ticker, "metric": "all"})
        if not data or "metric" not in data:
            return None

        m = data["metric"]
        result = {
            # 估值
            "pe_ttm": m.get("peTTM"),
            "pe_annual": m.get("peAnnual"),
            "pb_annual": m.get("pbAnnual"),
            "ps_ttm": m.get("psTTM"),
            "ev_ebitda": m.get("currentEv/freeCashFlowTTM"),
            "peg": m.get("pegRatio"),
            "dividend_yield": m.get("dividendYieldIndicatedAnnual"),

            # 盈利能力
            "roe_ttm": m.get("roeTTM"),
            "roa_ttm": m.get("roaTTM"),
            "roi_ttm": m.get("roiTTM"),
            "gross_margin_ttm": m.get("grossMarginTTM"),
            "operating_margin_ttm": m.get("operatingMarginTTM"),
            "net_margin_ttm": m.get("netProfitMarginTTM"),

            # 成長性
            "revenue_growth_3y": m.get("revenueGrowth3Y"),
            "revenue_growth_5y": m.get("revenueGrowth5Y"),
            "eps_growth_3y": m.get("epsGrowth3Y"),
            "eps_growth_5y": m.get("epsGrowth5Y"),
            "revenue_growth_qoq": m.get("revenueGrowthQuarterlyYoy"),
            "eps_growth_qoq": m.get("epsGrowthQuarterlyYoy"),

            # 財務健康
            "debt_equity": m.get("totalDebt/totalEquityAnnual"),
            "current_ratio": m.get("currentRatioAnnual"),
            "quick_ratio": m.get("quickRatioAnnual"),
            "fcf_per_share": m.get("freeCashFlowPerShareTTM"),

            # 價格相關
            "52w_high": m.get("52WeekHigh"),
            "52w_low": m.get("52WeekLow"),
            "52w_high_date": m.get("52WeekHighDate"),
            "52w_low_date": m.get("52WeekLowDate"),
            "beta": m.get("beta"),
            "10d_avg_vol": m.get("10DayAverageTradingVolume"),

            # 原始數據（備用）
            "_raw_count": len(m),
        }

        self.cache[cache_key] = result
        return result

    # ==================== 公司資料 ====================
    def get_company_profile(self, ticker):
        """獲取公司基本資料"""
        data = self._get("stock/profile2", {"symbol": ticker})
        if not data:
            return None
        return {
            "name": data.get("name"),
            "industry": data.get("finnhubIndustry"),
            "country": data.get("country"),
            "market_cap": data.get("marketCapitalization"),
            "ipo_date": data.get("ipo"),
            "exchange": data.get("exchange"),
            "web": data.get("weburl"),
        }

    # ==================== 新聞情緒 ====================
    def get_news_sentiment(self, ticker, days=7):
        """獲取近期新聞"""
        from_date = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
        to_date = datetime.now().strftime("%Y-%m-%d")
        data = self._get("company-news", {
            "symbol": ticker, "from": from_date, "to": to_date
        })
        if not data or not isinstance(data, list):
            return {"count": 0, "sentiment": "無數據"}

        return {
            "count": len(data),
            "headlines": [{"title": n.get("headline", ""), "source": n.get("source", "")}
                         for n in data[:5]],
        }

    # ==================== 綜合分析報告 ====================
    def full_analysis(self, ticker):
        """一鍵獲取某支股票的所有 Finnhub 數據"""
        print(f"  [Finnhub] 分析 {ticker}...", end=" ")

        profile = self.get_company_profile(ticker)
        ratings = self.get_analyst_ratings(ticker)
        target = self.get_price_target(ticker)
        insider = self.get_insider_transactions(ticker)
        metrics = self.get_financial_metrics(ticker)
        news = self.get_news_sentiment(ticker)

        # 綜合評分
        score = 0
        reasons = []

        # 分析師評級加分
        if ratings:
            if ratings["bull_pct"] > 70:
                score += 2
                reasons.append(f"分析師{ratings['bull_pct']}%看多({ratings['consensus']})")
            elif ratings["bull_pct"] > 50:
                score += 1
            elif ratings["bull_pct"] < 30:
                score -= 2
                reasons.append(f"分析師僅{ratings['bull_pct']}%看多")

        # 內部人交易信號
        if insider:
            if "看多" in insider["net_signal"]:
                score += 2
                reasons.append(f"內部人淨買入（買{insider['buys']}/賣{insider['sells']}）")
            elif "看空" in insider["net_signal"]:
                score -= 1
                reasons.append(f"內部人淨賣出（買{insider['buys']}/賣{insider['sells']}）")

        # 目標價上行空間
        if target and target.get("target_mean") and metrics:
            # 用52週高低推算當前價位
            pass  # 需要結合當前價格

        # 財務指標
        if metrics:
            roe = metrics.get("roe_ttm")
            if roe and roe > 20:
                score += 1
                reasons.append(f"ROE優秀({roe}%)")

            growth = metrics.get("revenue_growth_qoq")
            if growth and growth > 15:
                score += 1
                reasons.append(f"營收季增{growth}%")

        print(f"評分:{score}, 原因:{len(reasons)}個")

        return {
            "ticker": ticker,
            "profile": profile,
            "analyst_ratings": ratings,
            "price_target": target,
            "insider": insider,
            "metrics": metrics,
            "news": news,
            "finnhub_score": score,
            "finnhub_reasons": reasons,
        }
