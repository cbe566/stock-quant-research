"""
多線程篩選引擎 v2
=================
用 ThreadPoolExecutor 平行處理 3,775+ 隻股票
目標：15-20 分鐘完成全部四大市場篩選

架構：
  - 每市場獨立一個 thread pool
  - 每個 pool 內 10 個 worker 平行抓數據+評分
  - 自動重試失敗的股票
  - 進度條即時顯示
"""

import sys
import os
import json
import time
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "backtest_system"))

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# 偵測環境
IS_CLOUD = os.environ.get("GITHUB_ACTIONS") == "true" or os.environ.get("CI") == "true" or os.environ.get("ORACLE_CLOUD") == "true"

if IS_CLOUD:
    from cloud_data_engine import CloudDataEngine
    DataEngineClass = CloudDataEngine
else:
    try:
        from data_engine import DataEngine
        DataEngineClass = DataEngine
    except ImportError:
        from cloud_data_engine import CloudDataEngine
        DataEngineClass = CloudDataEngine

try:
    from config import STRATEGY_PARAMS
except ImportError:
    STRATEGY_PARAMS = {
        "multifactor": {
            "quality_roe_min": 0.10, "quality_debt_max": 0.70,
            "value_pe_max": 25, "value_pb_max": 5,
            "momentum_lookback": 126, "momentum_exclude_recent": 21,
        },
    }

from index_constituents import get_all_constituents

# ==================== 全域設定 ====================
MAX_WORKERS = 12          # 每市場最大平行線程
RETRY_COUNT = 1           # 失敗重試次數
RATE_LIMIT_DELAY = 0.3    # 每次請求間隔（秒）

# 線程安全鎖
_print_lock = threading.Lock()
_counter_lock = threading.Lock()


class MultiThreadScreener:
    """多線程全球股票篩選器"""

    def __init__(self):
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.start_date = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
        self.results = {}
        self.errors = []
        self._progress = {}

    def _create_engine(self):
        """每個線程建立自己的 DataEngine（避免共享狀態衝突）"""
        return DataEngineClass()

    def _score_stock(self, ticker):
        """對單隻股票進行六維度評分（線程安全）"""
        de = self._create_engine()

        try:
            result = {
                "ticker": ticker, "name": "", "sector": "",
                "current_price": None, "total_score": 0,
                "buy_score": 0, "sell_score": 0, "signals": [],
            }

            # 1. 基本面
            info = de.get_stock_info(ticker)
            if info:
                result["name"] = info.get("shortName", info.get("longName", ticker))
                result["sector"] = info.get("sector", "")

            # 2. QVM 多因子
            mf_total = 50
            mf = self._calc_multifactor(de, ticker, info)
            if mf:
                mf_total = mf["total_score"]
                result["quality"] = mf["quality_score"]
                result["value"] = mf["value_score"]
                result["momentum"] = mf["momentum_score"]
                result["f_score"] = mf["f_score"]

                if mf_total >= 70:
                    result["total_score"] += 3
                    result["signals"].append(f"QVM高分({mf_total})")
                elif mf_total >= 55:
                    result["total_score"] += 1
                elif mf_total < 35:
                    result["total_score"] -= 3
                    result["signals"].append(f"QVM低分({mf_total})")
                elif mf_total < 45:
                    result["total_score"] -= 1

            # 3. 技術面
            try:
                prices = de.get_prices(ticker, self.start_date, self.today)
                if not prices.empty and len(prices) > 50:
                    df = de.calc_all_technicals(prices)
                    close = df['Close']

                    # 簡化技術面信號計算
                    tech_score = 0.0
                    # 均線
                    if len(df) > 200 and 'SMA_20' in df.columns and 'SMA_50' in df.columns and 'SMA_200' in df.columns:
                        if df['SMA_20'].iloc[-1] > df['SMA_50'].iloc[-1]: tech_score += 15
                        else: tech_score -= 15
                        if df['SMA_50'].iloc[-1] > df['SMA_200'].iloc[-1]: tech_score += 15
                        else: tech_score -= 15
                        if close.iloc[-1] > df['SMA_200'].iloc[-1]: tech_score += 10
                        else: tech_score -= 10

                    # RSI
                    if 'RSI_14' in df.columns:
                        rsi = df['RSI_14'].iloc[-1]
                        if rsi < 30: tech_score += 20
                        elif rsi > 70: tech_score -= 20
                        elif rsi < 45: tech_score += 10
                        elif rsi > 55: tech_score -= 10

                    # MACD
                    if 'MACD' in df.columns and 'MACD_Signal' in df.columns:
                        if df['MACD'].iloc[-1] > df['MACD_Signal'].iloc[-1]: tech_score += 10
                        else: tech_score -= 10

                    tech_score = max(-100, min(100, tech_score))
                    result["tech_signal"] = round(tech_score, 1)

                    if tech_score > 40:
                        result["total_score"] += 2
                        result["signals"].append(f"技術面強勢({tech_score:+.0f})")
                    elif tech_score > 20:
                        result["total_score"] += 1
                    elif tech_score < -40:
                        result["total_score"] -= 2
                        result["signals"].append(f"技術面弱勢({tech_score:+.0f})")
                    elif tech_score < -20:
                        result["total_score"] -= 1

                    # 4. 均值回歸 Z-Score
                    if len(close) > 20:
                        mean20 = close.rolling(20).mean()
                        std20 = close.rolling(20).std()
                        z = float((close.iloc[-1] - mean20.iloc[-1]) / std20.iloc[-1]) if std20.iloc[-1] > 0 else 0
                        result["zscore"] = round(z, 2)

                        if z < -2.0:
                            result["total_score"] += 2
                            result["signals"].append(f"極度超賣(Z={z:.1f})")
                        elif z < -1.0:
                            result["total_score"] += 1
                            result["signals"].append(f"超賣(Z={z:.1f})")
                        elif z > 2.0:
                            result["total_score"] -= 2
                            result["signals"].append(f"極度超買(Z={z:.1f})")
                        elif z > 1.0:
                            result["total_score"] -= 1
                            result["signals"].append(f"超買(Z={z:.1f})")
            except Exception:
                pass

            # 5. F-Score
            f_score = mf.get("f_score", 5) if mf else 5
            result["f_score"] = f_score
            if f_score >= 7:
                result["total_score"] += 1
                result["signals"].append(f"F-Score高({f_score}/9)")
            elif f_score <= 3:
                result["total_score"] -= 1
                result["signals"].append(f"F-Score低({f_score}/9)")

            # 6. 分析師目標價
            if info:
                target = info.get("targetMeanPrice")
                current = info.get("currentPrice") or info.get("regularMarketPrice")
                if target and current and current > 0:
                    upside = round((target / current - 1) * 100, 1)
                    result["current_price"] = round(current, 2)
                    result["analyst_target"] = round(target, 2)
                    result["upside_pct"] = upside
                    if upside > 25:
                        result["total_score"] += 1
                        result["signals"].append(f"目標價上行{upside}%")
                    elif upside < -15:
                        result["total_score"] -= 1
                        result["signals"].append(f"目標價下行{upside}%")
                elif current:
                    result["current_price"] = round(current, 2)

            # 買賣分數
            s = result["total_score"]
            result["buy_score"] = max(0, min(100, 50 + s * 8))
            result["sell_score"] = max(0, min(100, 50 - s * 8))

            return result

        except Exception as e:
            return None

    def _calc_multifactor(self, de, ticker, info):
        """QVM 多因子評分"""
        if not info:
            return None

        scores = {}
        weights = {"quality": 0.35, "value": 0.35, "momentum": 0.30}

        # 質量因子
        q = 0
        roe = info.get("returnOnEquity")
        if roe is not None:
            if roe > 0.20: q += 2
            elif roe > 0.10: q += 1
        gm = info.get("grossMargins")
        if gm is not None:
            if gm > 0.40: q += 1.5
            elif gm > 0.25: q += 0.75
        dte = info.get("debtToEquity")
        if dte is not None:
            if dte < 50: q += 1
            elif dte < 100: q += 0.5
        fcf = info.get("freeCashflow")
        if fcf and fcf > 0: q += 0.5
        scores["quality"] = min(q / 5, 1.0) * 100

        # 價值因子
        v = 0
        pe = info.get("forwardPE") or info.get("trailingPE")
        if pe and 0 < pe < 15: v += 2
        elif pe and 0 < pe < 25: v += 1
        pb = info.get("priceToBook")
        if pb and 0 < pb < 2: v += 1.5
        elif pb and 0 < pb < 5: v += 0.75
        peg = info.get("pegRatio")
        if peg and 0 < peg < 1: v += 1
        elif peg and 0 < peg < 2: v += 0.5
        dy = info.get("dividendYield")
        if dy and dy > 0.02: v += 0.5
        scores["value"] = min(v / 5, 1.0) * 100

        # 動量因子
        m = 0
        try:
            prices = de.get_prices(ticker, self.start_date, self.today)
            if len(prices) > 126:
                close = prices['Close']
                mom = (close.iloc[-22] / close.iloc[-148]) - 1 if len(close) > 148 else 0
                if mom > 0.15: m += 1.5
                elif mom > 0.05: m += 0.75
                sma200 = close.rolling(200).mean()
                if len(sma200.dropna()) > 0 and close.iloc[-1] > sma200.iloc[-1]: m += 1
                rsi = de.calc_rsi(close, 14)
                if 40 < rsi.iloc[-1] < 70: m += 0.5
        except Exception:
            pass
        scores["momentum"] = min(m / 3, 1.0) * 100

        total = sum(scores.get(k, 0) * weights[k] for k in weights)

        # F-Score
        f_score = self._calc_f_score(info)

        return {
            "total_score": round(total, 1),
            "quality_score": round(scores.get("quality", 0), 1),
            "value_score": round(scores.get("value", 0), 1),
            "momentum_score": round(scores.get("momentum", 0), 1),
            "f_score": f_score,
        }

    def _calc_f_score(self, info):
        """Piotroski F-Score"""
        score = 0
        try:
            if info.get("returnOnAssets", 0) and info["returnOnAssets"] > 0: score += 1
            if info.get("freeCashflow", 0) and info["freeCashflow"] > 0: score += 1
            if info.get("earningsGrowth", 0) and info["earningsGrowth"] > 0: score += 1
            fcf = info.get("freeCashflow", 0)
            ni = info.get("netIncomeToCommon", 0)
            if fcf and ni and fcf > ni: score += 1
            if info.get("debtToEquity") is not None and info["debtToEquity"] < 100: score += 1
            if info.get("currentRatio", 0) and info["currentRatio"] > 1: score += 1
            if info.get("sharesOutstanding", 0) and info["sharesOutstanding"] > 0: score += 1
            if info.get("grossMargins", 0) and info["grossMargins"] > 0.3: score += 1
            if info.get("revenueGrowth", 0) and info["revenueGrowth"] > 0: score += 1
        except Exception:
            pass
        return score

    def screen_market(self, market_name, stocks):
        """多線程篩選一個市場"""
        total = len(stocks)
        completed = 0
        results = []
        errors = []

        print(f"\n{'='*60}")
        print(f"  {market_name}（{total} 隻股票，{MAX_WORKERS} 線程平行）")
        print(f"{'='*60}")

        start_time = time.time()

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_ticker = {executor.submit(self._score_stock, t): t for t in stocks}

            for future in as_completed(future_to_ticker):
                ticker = future_to_ticker[future]
                completed += 1

                try:
                    r = future.result(timeout=60)
                    if r:
                        results.append(r)
                        score = r['total_score']
                        # 每 50 隻輸出一次進度
                        if completed % 50 == 0 or completed == total:
                            elapsed = time.time() - start_time
                            speed = completed / elapsed if elapsed > 0 else 0
                            eta = (total - completed) / speed if speed > 0 else 0
                            with _print_lock:
                                print(f"  [{completed}/{total}] {speed:.1f} 隻/秒 | ETA {eta:.0f}秒 | 最新: {ticker} ({score:+d})")
                    else:
                        errors.append(ticker)
                except Exception as e:
                    errors.append(ticker)

        elapsed = time.time() - start_time
        results.sort(key=lambda x: x["total_score"], reverse=True)

        print(f"  完成：{len(results)}/{total} 成功 | {len(errors)} 失敗 | 耗時 {elapsed:.0f}秒")

        if errors:
            self.errors.extend([f"{market_name}/{t}" for t in errors[:10]])

        return results

    def run(self):
        """執行全部市場篩選"""
        total_start = time.time()

        print(f"\n{'#'*60}")
        print(f"  多線程全球股票篩選引擎 v2")
        print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'#'*60}")

        # 動態獲取成分股
        markets = get_all_constituents()

        for market_name, config in markets.items():
            stocks = config["stocks"]
            if stocks:
                self.results[market_name] = self.screen_market(market_name, stocks)

        total_elapsed = time.time() - total_start
        total_stocks = sum(len(v) for v in self.results.values())
        print(f"\n{'='*60}")
        print(f"  全部完成！{total_stocks} 隻股票 | 耗時 {total_elapsed/60:.1f} 分鐘")
        print(f"{'='*60}")

        return self.results

    def save(self):
        """儲存結果"""
        date_str = datetime.now().strftime("%Y%m%d")
        reports_dir = Path(__file__).parent / "daily_reports"
        reports_dir.mkdir(exist_ok=True)

        # JSON 原始數據
        json_path = reports_dir / f"screening_data_{date_str}.json"
        json_data = {
            "date": self.today,
            "generated_at": datetime.now().isoformat(),
            "markets": self.results,
            "errors": self.errors,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2, default=str)

        print(f"數據已儲存：{json_path}")
        return str(json_path)


def main():
    screener = MultiThreadScreener()
    screener.run()
    json_path = screener.save()
    return json_path


if __name__ == "__main__":
    main()
