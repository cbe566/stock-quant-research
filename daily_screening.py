"""
每日全球股票篩選系統
====================
每天自動執行 QVM 多因子 + 技術面 + 動量 + 均值回歸分析
覆蓋：美股、港股、台股、日股
每市場篩選：10 隻買入候選 + 10 隻賣出候選
輸出：Markdown 報告（存檔 + 顯示）

用法：
  python daily_screening.py          # 完整篩選
  python daily_screening.py --quick  # 快速模式（僅用緩存/縮小股池）
"""

import sys
import os
import json
import time
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# 確保可以 import backtest_system 模組
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "backtest_system"))

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# 偵測是否在 GitHub Actions 雲端環境
IS_CLOUD = os.environ.get("GITHUB_ACTIONS") == "true" or os.environ.get("CI") == "true"

if IS_CLOUD:
    # 雲端模式：使用多源備援引擎（不依賴 yfinance 單一來源）
    sys.path.insert(0, os.path.dirname(__file__))
    from cloud_data_engine import CloudDataEngine as DataEngine
    from cloud_data_engine import CloudDataEngine
    print("☁️ 雲端模式啟動 — 使用多源備援數據引擎")

    # 雲端環境下不需要 backtest_system 的 config
    class _CloudConfig:
        STRATEGY_PARAMS = {
            "multifactor": {
                "quality_roe_min": 0.10,
                "quality_debt_max": 0.70,
                "value_pe_max": 25,
                "value_pb_max": 5,
                "momentum_lookback": 126,
                "momentum_exclude_recent": 21,
                "rebalance_days": 63,
            },
        }
    STRATEGY_PARAMS = _CloudConfig.STRATEGY_PARAMS
else:
    # 本地模式：使用原有引擎
    import yfinance as yf
    from data_engine import DataEngine
    from config import STRATEGY_PARAMS

# 雲端模式下動態導入 StrategyEngine 或使用內建版本
if IS_CLOUD:
    # 雲端模式：策略引擎直接嵌入（避免 import 路徑問題）
    StrategyEngine = None  # 後面用 CloudStrategyEngine 替代
else:
    from strategies import StrategyEngine

# ==================== 全球股票池（動態抓取） ====================
from index_constituents import get_all_constituents

# 動態獲取四大市場指數成分股（雲端模式精簡股池）
ALL_MARKETS = get_all_constituents(cloud_mode=IS_CLOUD)


# ==================== 核心篩選引擎 ====================

class DailyScreener:
    """每日全球股票篩選器"""

    def __init__(self):
        if IS_CLOUD:
            self.de = CloudDataEngine()
        else:
            self.de = DataEngine()

        # 雲端模式下，策略引擎直接用 self（內建策略方法）
        # 本地模式下，用 backtest_system 的 StrategyEngine
        if IS_CLOUD:
            self.se = self  # 雲端模式：策略方法內建在 DailyScreener 中
        else:
            self.se = StrategyEngine(self.de)

        self.today = datetime.now().strftime("%Y-%m-%d")
        # 技術面需要 ~1 年數據
        self.start_date = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
        self.results = {}
        self.errors = []

    # === 雲端模式內建策略方法（與 StrategyEngine 接口兼容） ===

    def multifactor_score(self, ticker, params):
        """QVM 多因子評分 — 雲端模式用"""
        info = self.de.get_stock_info(ticker)
        if not info:
            return None

        scores = {}
        weights = {"quality": 0.35, "value": 0.35, "momentum": 0.30}

        # 質量因子
        q_score = 0
        q_max = 5
        roe = info.get("returnOnEquity")
        if roe is not None:
            if roe > 0.20: q_score += 2
            elif roe > 0.10: q_score += 1
        gm = info.get("grossMargins")
        if gm is not None:
            if gm > 0.40: q_score += 1.5
            elif gm > 0.25: q_score += 0.75
        dte = info.get("debtToEquity")
        if dte is not None:
            if dte < 50: q_score += 1
            elif dte < 100: q_score += 0.5
        fcf = info.get("freeCashflow")
        if fcf is not None and fcf > 0:
            q_score += 0.5
        scores["quality"] = min(q_score / q_max, 1.0) * 100

        # 價值因子
        v_score = 0
        v_max = 5
        pe = info.get("forwardPE") or info.get("trailingPE")
        if pe is not None:
            if 0 < pe < 15: v_score += 2
            elif 0 < pe < 25: v_score += 1
        pb = info.get("priceToBook")
        if pb is not None:
            if 0 < pb < 2: v_score += 1.5
            elif 0 < pb < 5: v_score += 0.75
        peg = info.get("pegRatio")
        if peg is not None:
            if 0 < peg < 1: v_score += 1
            elif 0 < peg < 2: v_score += 0.5
        dy = info.get("dividendYield")
        if dy is not None and dy > 0.02:
            v_score += 0.5
        scores["value"] = min(v_score / v_max, 1.0) * 100

        # 動量因子
        m_score = 0
        m_max = 3
        try:
            end = datetime.now().strftime("%Y-%m-%d")
            start = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
            prices = self.de.get_prices(ticker, start, end)
            if len(prices) > 126:
                close = prices['Close']
                mom_6m = (close.iloc[-22] / close.iloc[-148]) - 1 if len(close) > 148 else 0
                if mom_6m > 0.15: m_score += 1.5
                elif mom_6m > 0.05: m_score += 0.75
                sma200 = close.rolling(200).mean()
                if len(sma200.dropna()) > 0 and close.iloc[-1] > sma200.iloc[-1]:
                    m_score += 1
                rsi = self.de.calc_rsi(close, 14)
                if 40 < rsi.iloc[-1] < 70:
                    m_score += 0.5
        except Exception:
            pass
        scores["momentum"] = min(m_score / m_max, 1.0) * 100

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
            "recommendation": self._get_cloud_recommendation(total),
        }

    @staticmethod
    def _get_cloud_recommendation(score):
        if score >= 75: return "強烈買入"
        elif score >= 60: return "買入"
        elif score >= 45: return "觀望"
        elif score >= 30: return "減持"
        else: return "賣出"

    def technical_signals(self, ticker, start, end, params=None):
        """技術面綜合信號 — 雲端模式用"""
        prices = self.de.get_prices(ticker, start, end)
        if prices.empty:
            return pd.DataFrame()

        df = self.de.calc_all_technicals(prices)
        close = df['Close']

        signals = pd.DataFrame(index=df.index)
        signals['close'] = close

        # 均線
        ma_signal = pd.Series(0.0, index=df.index)
        ma_signal += np.where(df['SMA_20'] > df['SMA_50'], 15, -15)
        ma_signal += np.where(df['SMA_50'] > df['SMA_200'], 15, -15)
        ma_signal = ma_signal.astype(float)
        ma_signal += np.where(close > df['SMA_200'], 10, -10)
        signals['ma_signal'] = ma_signal.clip(-30, 30)

        # RSI
        rsi = df['RSI_14']
        rsi_signal = np.where(rsi < 30, 20, np.where(rsi > 70, -20,
                              np.where(rsi < 45, 10, np.where(rsi > 55, -10, 0))))
        signals['rsi_signal'] = pd.Series(rsi_signal, index=df.index).astype(float)

        # MACD
        macd_signal_val = pd.Series(0.0, index=df.index)
        macd_signal_val += np.where(df['MACD'] > df['MACD_Signal'], 10, -10)
        macd_signal_val += np.where(df['MACD_Hist'] > 0, 5, -5)
        macd_hist_dir = df['MACD_Hist'].diff()
        macd_signal_val += np.where(macd_hist_dir > 0, 5, -5)
        signals['macd_signal'] = pd.Series(macd_signal_val, index=df.index).clip(-20, 20)

        # 布林通道
        bb_signal = np.where(df['BB_PctB'] < 0.0, 15, np.where(df['BB_PctB'] > 1.0, -15,
                              np.where(df['BB_PctB'] < 0.2, 8, np.where(df['BB_PctB'] > 0.8, -8, 0))))
        signals['bb_signal'] = pd.Series(bb_signal, index=df.index).astype(float)

        # 成交量
        vol_signal = pd.Series(0.0, index=df.index)
        vol_ratio = df['VOL_RATIO']
        price_change = close.pct_change()
        vol_signal += np.where((vol_ratio > 1.5) & (price_change > 0.01), 10, 0)
        vol_signal += np.where((vol_ratio > 1.5) & (price_change < -0.01), -10, 0)
        signals['vol_signal'] = pd.Series(vol_signal, index=df.index).clip(-15, 15)

        signals['total_signal'] = (
            signals['ma_signal'] + signals['rsi_signal'] + signals['macd_signal'] +
            signals['bb_signal'] + signals['vol_signal']
        ).clip(-100, 100)

        signals['direction'] = np.where(signals['total_signal'] > 25, "買入",
                               np.where(signals['total_signal'] < -25, "賣出", "觀望"))
        return signals

    def mean_reversion_signals(self, ticker, start, end, lookback=20):
        """均值回歸信號 — 雲端模式用"""
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

    def screen_single_stock(self, ticker):
        """
        對單隻股票執行完整六維度分析
        回傳標準化的評分結果 dict，失敗回傳 None
        """
        try:
            result = {
                "ticker": ticker,
                "name": "",
                "sector": "",
                "current_price": None,
                "total_score": 0,        # 綜合得分（-10 ~ +10）
                "buy_score": 0,           # 買入信號強度（0~100）
                "sell_score": 0,          # 賣出信號強度（0~100）
                "signals": [],            # 信號理由列表
            }

            # --- 0. 預先抓取價格（避免後續重複呼叫） ---
            if IS_CLOUD:
                try:
                    prices = self.de.get_prices(ticker, self.start_date, self.today)
                    if prices is not None and not prices.empty:
                        # 存入快取，後續 multifactor / technical / mean_reversion 自動命中
                        cache_key = f"{ticker}_{self.start_date}_{self.today}"
                        self.de.price_cache[cache_key] = prices
                except Exception:
                    pass

            # --- 1. 基本面 ---
            info = self.de.get_stock_info(ticker)
            if info:
                result["name"] = info.get("shortName", info.get("longName", ticker))
                result["sector"] = info.get("sector", "")

            # --- 2. QVM 多因子評分 ---
            mf = self.se.multifactor_score(ticker, STRATEGY_PARAMS["multifactor"])
            mf_total = 50  # 預設中性
            if mf:
                mf_total = mf["total_score"]
                result["quality"] = mf["quality_score"]
                result["value"] = mf["value_score"]
                result["momentum"] = mf["momentum_score"]
                result["f_score"] = mf["f_score"]
                result["recommendation"] = mf["recommendation"]

                # 多因子信號
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

            # --- 3. 技術面信號 ---
            try:
                tech = self.se.technical_signals(ticker, self.start_date, self.today)
                if not tech.empty and len(tech) > 5:
                    latest = float(tech['total_signal'].iloc[-1])
                    avg_5d = float(tech['total_signal'].iloc[-5:].mean())
                    result["tech_signal"] = round(latest, 1)
                    result["tech_5d"] = round(avg_5d, 1)
                    result["tech_direction"] = tech['direction'].iloc[-1]

                    if latest > 40:
                        result["total_score"] += 2
                        result["signals"].append(f"技術面強勢({latest:+.0f})")
                    elif latest > 20:
                        result["total_score"] += 1
                    elif latest < -40:
                        result["total_score"] -= 2
                        result["signals"].append(f"技術面弱勢({latest:+.0f})")
                    elif latest < -20:
                        result["total_score"] -= 1

                    # 趨勢轉折（5日均值 vs 最新）加分
                    if latest > 20 and avg_5d < 0:
                        result["total_score"] += 1
                        result["signals"].append("技術面轉多")
                    elif latest < -20 and avg_5d > 0:
                        result["total_score"] -= 1
                        result["signals"].append("技術面轉空")
            except Exception:
                pass

            # --- 4. 均值回歸 ---
            try:
                mr = self.se.mean_reversion_signals(ticker, self.start_date, self.today)
                if not mr.empty:
                    z = float(mr['zscore'].iloc[-1])
                    result["zscore"] = round(z, 2)
                    result["mr_signal"] = mr['signal'].iloc[-1]

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

            # --- 5. F-Score ---
            f_score = result.get("f_score", 5)
            if f_score >= 7:
                result["total_score"] += 1
                result["signals"].append(f"F-Score高({f_score}/9)")
            elif f_score <= 3:
                result["total_score"] -= 1
                result["signals"].append(f"F-Score低({f_score}/9)")

            # --- 6. 分析師目標價 ---
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

            # --- 計算買入/賣出分數 ---
            s = result["total_score"]
            result["buy_score"] = max(0, min(100, 50 + s * 8))
            result["sell_score"] = max(0, min(100, 50 - s * 8))

            return result

        except Exception as e:
            self.errors.append(f"{ticker}: {str(e)}")
            return None

    def screen_market(self, market_name, stocks, time_budget=None):
        """
        篩選整個市場，回傳排序後的結果
        time_budget: 該市場的時間預算（秒），超時則提前結束
        """
        print(f"\n{'='*60}")
        print(f"  正在篩選 {market_name}（{len(stocks)} 隻股票）...")
        print(f"{'='*60}")

        results = []
        total = len(stocks)
        market_start = time.time()
        completed_count = 0
        failed_count = 0

        if IS_CLOUD:
            # 雲端模式：多線程並行（8 個線程）
            max_workers = 8
            print(f"  ⚡ 多線程模式（{max_workers} 線程並行）")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_ticker = {
                    executor.submit(self.screen_single_stock, ticker): ticker
                    for ticker in stocks
                }

                for future in as_completed(future_to_ticker):
                    ticker = future_to_ticker[future]
                    completed_count += 1

                    # 時間預算檢查
                    if time_budget and (time.time() - market_start) > time_budget:
                        print(f"\n  ⏰ {market_name} 時間預算用盡，已完成 {completed_count}/{total}")
                        # 取消未完成的任務
                        for f in future_to_ticker:
                            f.cancel()
                        break

                    try:
                        r = future.result(timeout=30)
                        if r:
                            results.append(r)
                            score = r['total_score']
                            indicator = "🟢" if score >= 3 else ("🔴" if score <= -3 else "⚪")
                            print(f"  [{completed_count}/{total}] {ticker}... {indicator} 得分:{score:+d}  {r.get('name', '')[:20]}")
                        else:
                            failed_count += 1
                    except Exception:
                        failed_count += 1

                    # 每 50 隻輸出進度
                    if completed_count % 50 == 0:
                        elapsed = time.time() - market_start
                        speed = completed_count / elapsed if elapsed > 0 else 0
                        remaining = (total - completed_count) / speed if speed > 0 else 0
                        print(f"  📊 進度 {completed_count}/{total} | 速度 {speed:.1f}隻/秒 | 預估剩餘 {remaining:.0f}秒")
        else:
            # 本地模式：序列執行（原始邏輯）
            for i, ticker in enumerate(stocks, 1):
                print(f"  [{i}/{total}] {ticker}...", end=" ", flush=True)
                r = self.screen_single_stock(ticker)
                if r:
                    results.append(r)
                    score = r['total_score']
                    indicator = "🟢" if score >= 3 else ("🔴" if score <= -3 else "⚪")
                    print(f"{indicator} 得分:{score:+d}  {r.get('name', '')[:20]}")
                else:
                    print("❌ 失敗")
                    failed_count += 1

                if i % 10 == 0:
                    time.sleep(1)

        elapsed = time.time() - market_start
        print(f"  ✅ {market_name} 完成：{len(results)} 成功 / {failed_count} 失敗 / 耗時 {elapsed:.0f}秒")

        # 排序
        results.sort(key=lambda x: x["total_score"], reverse=True)
        return results

    def run_full_screening(self):
        """執行全部四大市場篩選"""
        start_time = time.time()
        # 雲端模式總時間預算 22 分鐘（留 3 分鐘給報告生成和 git push）
        TOTAL_BUDGET = 22 * 60 if IS_CLOUD else None

        print(f"\n{'#'*60}")
        print(f"  每日全球股票篩選系統")
        print(f"  執行時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if IS_CLOUD:
            total_stocks = sum(len(c["stocks"]) for c in ALL_MARKETS.values())
            print(f"  ☁️ 雲端模式 | 總股數：{total_stocks} | 時間預算：22分鐘")
        print(f"{'#'*60}")

        # 按股票數量比例分配時間預算
        total_stocks = sum(len(c["stocks"]) for c in ALL_MARKETS.values())

        for market_name, config in ALL_MARKETS.items():
            # 計算該市場的時間預算
            market_budget = None
            if TOTAL_BUDGET:
                elapsed = time.time() - start_time
                remaining = TOTAL_BUDGET - elapsed
                if remaining <= 60:
                    print(f"\n  ⏰ 總時間預算即將用盡，跳過 {market_name}")
                    continue
                # 按剩餘市場的股票比例分配
                remaining_stocks = sum(
                    len(c["stocks"]) for mn, c in ALL_MARKETS.items()
                    if mn not in self.results
                )
                ratio = len(config["stocks"]) / remaining_stocks if remaining_stocks > 0 else 1
                market_budget = remaining * ratio

            try:
                results = self.screen_market(market_name, config["stocks"], time_budget=market_budget)
                self.results[market_name] = results
            except Exception as e:
                print(f"\n  ⚠️ {market_name} 篩選出錯：{e}")
                self.errors.append(f"{market_name}: {str(e)}")

        elapsed = time.time() - start_time
        print(f"\n✅ 全部篩選完成，耗時 {elapsed/60:.1f} 分鐘")
        if self.errors:
            print(f"⚠️ 有 {len(self.errors)} 個錯誤")

        # 雲端模式：輸出數據源統計
        if IS_CLOUD and hasattr(self.de, 'print_source_stats'):
            self.de.print_source_stats()

        return self.results

    def generate_report(self):
        """生成 Markdown 報告"""
        today_str = datetime.now().strftime("%Y年%m月%d日")
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        weekday = weekdays[datetime.now().weekday()]

        lines = []
        lines.append(f"# 每日全球股票篩選報告")
        lines.append(f"")
        lines.append(f"**日期**：{today_str}（星期{weekday}）")
        lines.append(f"**生成時間**：{datetime.now().strftime('%H:%M:%S')}")
        lines.append(f"**篩選模型**：QVM多因子 + 技術面 + 動量 + 均值回歸 + F-Score + 分析師目標價")
        lines.append(f"")
        lines.append(f"---")

        for market_name, results in self.results.items():
            if not results:
                lines.append(f"\n## {market_name}")
                lines.append(f"\n> ⚠️ 本市場無有效篩選結果\n")
                continue

            lines.append(f"\n## {market_name}")
            lines.append(f"")

            # === 買入候選 TOP 10 ===
            buy_list = [r for r in results if r["total_score"] > 0]
            buy_top = buy_list[:10] if buy_list else results[:10]

            lines.append(f"### 🟢 買入關注（TOP 10）")
            lines.append(f"")
            lines.append(f"| 排名 | 代碼 | 名稱 | 現價 | 綜合得分 | QVM | 技術面 | Z-Score | 信號摘要 |")
            lines.append(f"|:---:|:---:|:---|:---:|:---:|:---:|:---:|:---:|:---|")

            for i, r in enumerate(buy_top[:10], 1):
                ticker = r["ticker"]
                name = (r.get("name", "") or "")[:15]
                price = r.get("current_price", "—")
                if isinstance(price, (int, float)):
                    price = f"{price:,.2f}"
                score = r["total_score"]
                qvm = r.get("quality", "—")
                if isinstance(qvm, (int, float)):
                    qvm = f"{qvm:.0f}"
                tech = r.get("tech_signal", "—")
                if isinstance(tech, (int, float)):
                    tech = f"{tech:+.0f}"
                z = r.get("zscore", "—")
                if isinstance(z, (int, float)):
                    z = f"{z:.1f}"
                signals = "、".join(r.get("signals", [])[:3]) or "—"
                lines.append(f"| {i} | **{ticker}** | {name} | {price} | **{score:+d}** | {qvm} | {tech} | {z} | {signals} |")

            lines.append(f"")

            # === 賣出候選 TOP 10 ===
            sell_list = sorted(results, key=lambda x: x["total_score"])
            sell_list = [r for r in sell_list if r["total_score"] < 0]
            sell_top = sell_list[:10] if sell_list else sorted(results, key=lambda x: x["total_score"])[:10]

            lines.append(f"### 🔴 賣出關注（TOP 10）")
            lines.append(f"")
            lines.append(f"| 排名 | 代碼 | 名稱 | 現價 | 綜合得分 | QVM | 技術面 | Z-Score | 信號摘要 |")
            lines.append(f"|:---:|:---:|:---|:---:|:---:|:---:|:---:|:---:|:---|")

            for i, r in enumerate(sell_top[:10], 1):
                ticker = r["ticker"]
                name = (r.get("name", "") or "")[:15]
                price = r.get("current_price", "—")
                if isinstance(price, (int, float)):
                    price = f"{price:,.2f}"
                score = r["total_score"]
                qvm = r.get("quality", "—")
                if isinstance(qvm, (int, float)):
                    qvm = f"{qvm:.0f}"
                tech = r.get("tech_signal", "—")
                if isinstance(tech, (int, float)):
                    tech = f"{tech:+.0f}"
                z = r.get("zscore", "—")
                if isinstance(z, (int, float)):
                    z = f"{z:.1f}"
                signals = "、".join(r.get("signals", [])[:3]) or "—"
                lines.append(f"| {i} | **{ticker}** | {name} | {price} | **{score:+d}** | {qvm} | {tech} | {z} | {signals} |")

            lines.append(f"")

            # === 市場概覽統計 ===
            total_stocks = len(results)
            bullish = len([r for r in results if r["total_score"] >= 3])
            bearish = len([r for r in results if r["total_score"] <= -3])
            neutral = total_stocks - bullish - bearish
            avg_score = sum(r["total_score"] for r in results) / total_stocks if total_stocks > 0 else 0

            lines.append(f"**篩選 {total_stocks} 隻** ｜ 看多 {bullish} · 看空 {bearish} ｜ 平均得分 {avg_score:+.1f}")
            lines.append(f"")
            lines.append(f"---")

        # === 風險提醒 ===
        lines.append(f"\n## ⚠️ 風險提醒")
        lines.append(f"")
        lines.append(f"- 本報告為量化模型自動生成，僅供研究參考，不構成投資建議")
        lines.append(f"- 買入/賣出信號基於歷史數據回溯，不保證未來表現")
        lines.append(f"- 投資前請結合宏觀環境、個股基本面、市場情緒做綜合判斷")
        lines.append(f"- 嚴格執行風控：單股不超過10%、組合止損15%、個股止損8%")
        lines.append(f"")

        # === 錯誤日誌 ===
        if self.errors:
            lines.append(f"\n## 📋 執行日誌")
            lines.append(f"")
            lines.append(f"以下股票篩選時出現問題（共 {len(self.errors)} 個）：")
            for err in self.errors[:20]:
                lines.append(f"- {err}")
            lines.append(f"")

        lines.append(f"\n---")
        lines.append(f"*報告由 QVM 量化篩選系統自動生成 | {datetime.now().strftime('%Y-%m-%d %H:%M')}*")

        return "\n".join(lines)

    def save_report(self):
        """儲存報告到檔案"""
        report = self.generate_report()
        date_str = datetime.now().strftime("%Y%m%d")

        # 儲存 Markdown 報告
        report_dir = Path(__file__).parent / "daily_reports"
        report_dir.mkdir(exist_ok=True)

        md_path = report_dir / f"每日篩選報告_{date_str}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"\n📄 報告已儲存：{md_path}")

        # 同時儲存 JSON 原始數據（方便後續分析）
        json_path = report_dir / f"screening_data_{date_str}.json"
        json_data = {
            "date": self.today,
            "generated_at": datetime.now().isoformat(),
            "markets": {},
            "errors": self.errors,
        }
        for market, results in self.results.items():
            json_data["markets"][market] = results

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2, default=str)
        print(f"📊 數據已儲存：{json_path}")

        # 也在根目錄放一份「最新報告」方便查看
        latest_path = Path(__file__).parent / "每日篩選報告_最新.md"
        with open(latest_path, "w", encoding="utf-8") as f:
            f.write(report)

        return str(md_path), report


# ==================== 主程式入口 ====================

def main():
    """主程式"""
    screener = DailyScreener()

    # 檢查是否是交易日（簡單判斷：週六日跳過）
    weekday = datetime.now().weekday()
    if weekday >= 5:
        print(f"⚠️ 今天是週{'六' if weekday == 5 else '日'}，非交易日")
        print("仍然執行篩選（使用最新收盤數據）...")

    # 執行篩選
    screener.run_full_screening()

    # 生成並儲存報告
    md_path, report = screener.save_report()

    # 輸出報告內容
    print(f"\n{'='*60}")
    print(report)

    return md_path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="每日全球股票篩選系統")
    parser.add_argument("--market", type=str, default=None,
                        help="只篩選指定市場（美股/港股/台股/日股）")
    parser.add_argument("--quick", action="store_true",
                        help="快速模式（僅用緩存/縮小股池）")
    parser.add_argument("--output-json", type=str, default=None,
                        help="指定 JSON 結果輸出路徑（用於 matrix job 合併）")
    args = parser.parse_args()

    if args.market:
        # 單市場模式：只篩選指定市場，輸出 JSON
        print(f"🎯 單市場模式：{args.market}")
        screener = DailyScreener()

        if args.market not in ALL_MARKETS:
            print(f"❌ 未知市場：{args.market}，可用：{list(ALL_MARKETS.keys())}")
            sys.exit(1)

        config = ALL_MARKETS[args.market]
        results = screener.screen_market(args.market, config["stocks"])
        screener.results[args.market] = results

        # 輸出 JSON 結果
        output_path = args.output_json or f"screening_{args.market}.json"
        json_data = {
            "date": screener.today,
            "generated_at": datetime.now().isoformat(),
            "market": args.market,
            "results": results,
            "errors": screener.errors,
        }
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2, default=str)
        print(f"✅ {args.market} 篩選完成，結果已存到 {output_path}")
    else:
        # 完整模式：四大市場全部篩選
        main()
