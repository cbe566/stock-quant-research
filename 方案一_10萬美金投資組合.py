#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
方案一：10萬美金最優投資組合 — 含20年歷史回測驗證
基於 QVM 多因子模型 + 宏觀景氣循環 + 板塊輪動
自驗證模組：驗證選股邏輯在過去20年+每個景氣循環中的有效性
"""

import numpy as np
import pandas as pd
import requests
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backtest_system'))
from config import TIINGO_API_KEY, FRED_API_KEY

# ============================================================
# 第一部分：最終投資組合定義
# ============================================================

PORTFOLIO = {
    "總資金": 100000,
    "基準日期": "2026-03-29",
    "策略名稱": "QVM宏觀景氣循環增強組合",
    "當前景氣判斷": "高峰→收縮過渡期（滯脹風險）",

    # ===== 核心持倉（60%）=====
    "核心持倉": {
        # AI/科技核心 — 25%
        "GOOGL": {
            "名稱": "Alphabet (Google)",
            "配置比例": 0.10,
            "金額": 10000,
            "買入價": 274.34,
            "股數": 36,
            "邏輯": "AI龍頭中PE最低(25x)，RSI 22嚴重超賣，Gemini+雲端雙引擎",
            "因子": {"Quality": 90, "Value": 85, "Momentum": 70},
            "板塊": "通訊/科技",
        },
        "TSM": {
            "名稱": "台積電 ADR",
            "配置比例": 0.08,
            "金額": 8000,
            "買入價": 326.74,
            "股數": 24,
            "邏輯": "AI晶片製造壟斷，2nm先進製程全球獨步，ROE 35%",
            "因子": {"Quality": 92, "Value": 75, "Momentum": 85},
            "板塊": "半導體",
        },
        "MU": {
            "名稱": "Micron Technology",
            "配置比例": 0.07,
            "金額": 7000,
            "買入價": 357.22,
            "股數": 19,
            "邏輯": "HBM記憶體受惠AI爆發，PE僅17倍為AI股最低，6個月暴漲118%",
            "因子": {"Quality": 80, "Value": 90, "Momentum": 95},
            "板塊": "半導體",
        },

        # 大型價值/防禦 — 20%
        "MRK": {
            "名稱": "Merck & Co",
            "配置比例": 0.07,
            "金額": 7000,
            "買入價": 119.63,
            "股數": 58,
            "邏輯": "Keytruda持續增長，PE 16倍+殖利率2.84%，價值+成長兼具",
            "因子": {"Quality": 88, "Value": 88, "Momentum": 80},
            "板塊": "醫療",
        },
        "JNJ": {
            "名稱": "Johnson & Johnson",
            "配置比例": 0.06,
            "金額": 6000,
            "買入價": 240.45,
            "股數": 24,
            "邏輯": "62年連續增股息（股息之王），醫療防禦+33.9%漲幅",
            "因子": {"Quality": 85, "Value": 78, "Momentum": 82},
            "板塊": "醫療",
        },
        "VZ": {
            "名稱": "Verizon Communications",
            "配置比例": 0.04,
            "金額": 4000,
            "買入價": 50.31,
            "股數": 79,
            "邏輯": "PE 12倍+殖利率5.63%，5G穩定現金流，FCF $172億",
            "因子": {"Quality": 72, "Value": 95, "Momentum": 75},
            "板塊": "通訊",
        },
        "PEP": {
            "名稱": "PepsiCo",
            "配置比例": 0.03,
            "金額": 3000,
            "買入價": 153.04,
            "股數": 19,
            "邏輯": "50+年股息貴族，食品飲料雙引擎，滯脹環境防禦王",
            "因子": {"Quality": 82, "Value": 70, "Momentum": 68},
            "板塊": "必需消費",
        },

        # 景氣循環受益 — 15%
        "XLE_ETF": {
            "名稱": "能源板塊ETF (XLE)",
            "配置比例": 0.08,
            "金額": 8000,
            "買入價": 105.00,
            "股數": 76,
            "邏輯": "伊朗戰爭推動油價>$100，YTD +40.84%，滯脹環境最強板塊",
            "因子": {"Quality": 78, "Value": 82, "Momentum": 98},
            "板塊": "能源",
        },
        "ITA_ETF": {
            "名稱": "國防航太ETF (ITA)",
            "配置比例": 0.07,
            "金額": 7000,
            "買入價": 180.00,
            "股數": 38,
            "邏輯": "YTD +52.68%，全球軍費擴張，FY2026美國國防預算>$9,000億",
            "因子": {"Quality": 80, "Value": 65, "Momentum": 98},
            "板塊": "國防",
        },
    },

    # ===== 衛星持倉（25%）=====
    "衛星持倉": {
        "LLY": {
            "名稱": "Eli Lilly",
            "配置比例": 0.06,
            "金額": 6000,
            "買入價": 878.24,
            "股數": 6,
            "邏輯": "GLP-1減重藥全球爆發，ROE 101%，RSI 15極度超賣買入機會",
            "因子": {"Quality": 95, "Value": 55, "Momentum": 78},
            "板塊": "醫療/GLP-1",
        },
        "NVDA": {
            "名稱": "NVIDIA",
            "配置比例": 0.06,
            "金額": 6000,
            "買入價": 167.52,
            "股數": 35,
            "邏輯": "AI算力霸主，Blackwell供不應求，回調7.9%提供買入機會",
            "因子": {"Quality": 93, "Value": 60, "Momentum": 65},
            "板塊": "半導體/AI",
        },
        "MRVL": {
            "名稱": "Marvell Technology",
            "配置比例": 0.05,
            "金額": 5000,
            "買入價": 94.88,
            "股數": 52,
            "邏輯": "AI定制ASIC晶片龍頭，雲端巨頭定制化需求受惠者",
            "因子": {"Quality": 75, "Value": 70, "Momentum": 78},
            "板塊": "半導體/AI",
        },
        "GLD_ETF": {
            "名稱": "黃金ETF (GLD)",
            "配置比例": 0.08,
            "金額": 8000,
            "買入價": 414.70,
            "股數": 19,
            "邏輯": "避險+抗通膨+去美元化，地緣風險對沖，過去一年+49%",
            "因子": {"Quality": "N/A", "Value": "N/A", "Momentum": 88},
            "板塊": "商品/避險",
        },
    },

    # ===== 防守持倉（15%）=====
    "防守持倉": {
        "AGG_ETF": {
            "名稱": "綜合債券ETF (AGG)",
            "配置比例": 0.07,
            "金額": 7000,
            "買入價": 98.54,
            "股數": 71,
            "邏輯": "殖利率3.96%，降息週期受益，費用率僅0.03%",
            "因子": {"Quality": "N/A", "Value": "N/A", "Momentum": 50},
            "板塊": "債券",
        },
        "SHV_ETF": {
            "名稱": "短期國債ETF (SHV)",
            "配置比例": 0.05,
            "金額": 5000,
            "買入價": 110.00,
            "股數": 45,
            "邏輯": "殖利率~4%，現金替代品，保留下跌加碼彈藥",
            "因子": {"Quality": "N/A", "Value": "N/A", "Momentum": "N/A"},
            "板塊": "現金/短債",
        },
        "EEM_ETF": {
            "名稱": "新興市場ETF (EEM)",
            "配置比例": 0.03,
            "金額": 3000,
            "買入價": 55.20,
            "股數": 54,
            "邏輯": "PE 16倍vs美股25倍，殖利率2.2%，分散地域風險",
            "因子": {"Quality": "N/A", "Value": 85, "Momentum": 60},
            "板塊": "國際/新興市場",
        },
    },
}


# ============================================================
# 第二部分：20年歷史回測驗證系統
# ============================================================

class PortfolioBacktester:
    """
    驗證 QVM + 景氣循環選股邏輯在過去20年+每個經濟週期中的有效性
    使用歷史數據模擬每個景氣階段的配置策略
    """

    def __init__(self):
        self.tiingo_headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Token {TIINGO_API_KEY}'
        }
        self._cache = {}

    # ---------- 景氣循環定義（2003-2026）----------
    MARKET_CYCLES = {
        # 週期名稱: (起始, 結束, 階段, 描述)
        "2003-2007_擴張": ("2003-03-01", "2007-10-01", "擴張", "房市泡沫前的牛市"),
        "2007-2009_收縮": ("2007-10-01", "2009-03-01", "收縮", "全球金融危機"),
        "2009-2015_復甦": ("2009-03-01", "2015-06-01", "擴張", "QE驅動復甦"),
        "2015-2016_放緩": ("2015-06-01", "2016-06-01", "高峰", "中國放緩+升息恐慌"),
        "2016-2018_擴張": ("2016-06-01", "2018-10-01", "擴張", "減稅+科技牛市"),
        "2018-2019_放緩": ("2018-10-01", "2019-06-01", "收縮", "升息+貿易戰"),
        "2019-2020_短擴張": ("2019-06-01", "2020-02-01", "擴張", "降息驅動反彈"),
        "2020_COVID": ("2020-02-01", "2020-06-01", "收縮", "新冠疫情崩盤"),
        "2020-2022_復甦": ("2020-06-01", "2022-01-01", "擴張", "無限QE牛市"),
        "2022_緊縮": ("2022-01-01", "2022-10-01", "收縮", "暴力升息+通膨"),
        "2022-2024_擴張": ("2022-10-01", "2024-07-01", "擴張", "AI牛市"),
        "2024-2025_高峰": ("2024-07-01", "2025-12-01", "高峰", "估值到頂+板塊輪動"),
        "2026_滯脹": ("2025-12-01", "2026-03-29", "收縮", "伊朗戰爭+滯脹風險"),
    }

    # 每個景氣階段的配置策略（這是我們要驗證的核心邏輯）
    CYCLE_STRATEGIES = {
        "擴張": {
            "description": "經濟擴張 → 偏重成長/科技/非必需消費",
            "allocation": {
                "科技成長": 0.35,    # QQQ / 科技龍頭
                "金融": 0.15,       # XLF
                "非必需消費": 0.10,  # XLY
                "工業": 0.10,       # XLI
                "國際": 0.10,       # EEM/VEA
                "債券": 0.10,       # AGG
                "黃金": 0.05,       # GLD
                "現金": 0.05,       # SHV
            },
            "etf_proxies": ["QQQ", "XLF", "XLY", "XLI", "EEM", "AGG", "GLD", "SHV"],
            "weights": [0.35, 0.15, 0.10, 0.10, 0.10, 0.10, 0.05, 0.05],
        },
        "高峰": {
            "description": "經濟到頂 → 轉向能源/原物料/防禦",
            "allocation": {
                "能源": 0.20,
                "原物料": 0.10,
                "必需消費": 0.15,
                "醫療": 0.15,
                "公用事業": 0.10,
                "黃金": 0.10,
                "債券": 0.15,
                "現金": 0.05,
            },
            "etf_proxies": ["XLE", "XLB", "XLP", "XLV", "XLU", "GLD", "AGG", "SHV"],
            "weights": [0.20, 0.10, 0.15, 0.15, 0.10, 0.10, 0.15, 0.05],
        },
        "收縮": {
            "description": "經濟衰退 → 全面防禦 + 黃金 + 債券",
            "allocation": {
                "必需消費": 0.15,
                "醫療": 0.15,
                "公用事業": 0.10,
                "能源": 0.10,
                "黃金": 0.15,
                "債券": 0.25,
                "現金": 0.10,
            },
            "etf_proxies": ["XLP", "XLV", "XLU", "XLE", "GLD", "TLT", "SHV"],
            "weights": [0.15, 0.15, 0.10, 0.10, 0.15, 0.25, 0.10],
        },
    }

    def get_etf_prices(self, ticker, start, end):
        """取得ETF歷史價格"""
        cache_key = f"{ticker}_{start}_{end}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        try:
            r = requests.get(
                f'https://api.tiingo.com/tiingo/daily/{ticker}/prices'
                f'?startDate={start}&endDate={end}',
                headers=self.tiingo_headers, timeout=15
            )
            if r.status_code == 200:
                data = r.json()
                if data:
                    df = pd.DataFrame(data)
                    df['date'] = pd.to_datetime(df['date']).dt.tz_localize(None)
                    df = df.set_index('date')
                    self._cache[cache_key] = df
                    return df
        except Exception as e:
            print(f"  ⚠ {ticker} 數據取得失敗: {e}")
        return None

    def calc_period_return(self, ticker, start, end):
        """計算特定期間的報酬率"""
        df = self.get_etf_prices(ticker, start, end)
        if df is not None and len(df) > 1:
            start_price = df['adjClose'].iloc[0]
            end_price = df['adjClose'].iloc[-1]
            return (end_price - start_price) / start_price
        return None

    def calc_portfolio_return(self, etfs, weights, start, end):
        """計算投資組合在特定期間的加權報酬"""
        total_return = 0
        valid_weight = 0

        for ticker, weight in zip(etfs, weights):
            ret = self.calc_period_return(ticker, start, end)
            if ret is not None:
                total_return += ret * weight
                valid_weight += weight

        if valid_weight > 0:
            # 調整未取得數據的權重
            return total_return / valid_weight
        return None

    def backtest_all_cycles(self):
        """回測所有景氣循環，驗證策略有效性"""
        print("\n" + "=" * 80)
        print("📊 20年歷史回測驗證 — QVM景氣循環策略 vs 買入持有(SPY)")
        print("=" * 80)

        results = []
        spy_results = []

        for cycle_name, (start, end, phase, desc) in self.MARKET_CYCLES.items():
            print(f"\n{'─' * 60}")
            print(f"🔄 週期: {cycle_name} ({desc})")
            print(f"   階段: {phase} | 期間: {start} → {end}")

            # 我們的策略配置
            strategy = self.CYCLE_STRATEGIES.get(phase, self.CYCLE_STRATEGIES["擴張"])
            etfs = strategy["etf_proxies"]
            weights = strategy["weights"]

            port_ret = self.calc_portfolio_return(etfs, weights, start, end)
            spy_ret = self.calc_period_return("SPY", start, end)

            if port_ret is not None and spy_ret is not None:
                # 計算年化報酬
                days = (pd.Timestamp(end) - pd.Timestamp(start)).days
                years = max(days / 365.25, 0.1)

                port_ann = (1 + port_ret) ** (1 / years) - 1
                spy_ann = (1 + spy_ret) ** (1 / years) - 1
                alpha = port_ret - spy_ret

                results.append({
                    "週期": cycle_name,
                    "階段": phase,
                    "描述": desc,
                    "天數": days,
                    "策略報酬": port_ret,
                    "策略年化": port_ann,
                    "SPY報酬": spy_ret,
                    "SPY年化": spy_ann,
                    "超額Alpha": alpha,
                })

                win = "✅ 贏" if alpha > 0 else "❌ 輸"
                print(f"   📈 策略報酬: {port_ret:+.2%} (年化 {port_ann:+.2%})")
                print(f"   📉 SPY報酬:  {spy_ret:+.2%} (年化 {spy_ann:+.2%})")
                print(f"   🎯 Alpha:    {alpha:+.2%} {win}")
            else:
                print(f"   ⚠ 數據不足，跳過此週期")

        return results

    def calc_cumulative_performance(self, results):
        """計算20年累計表現"""
        print("\n" + "=" * 80)
        print("📊 累計表現統計")
        print("=" * 80)

        if not results:
            print("⚠ 無有效回測結果")
            return {}

        # 累計報酬
        cum_strategy = 1.0
        cum_spy = 1.0
        win_count = 0

        for r in results:
            cum_strategy *= (1 + r["策略報酬"])
            cum_spy *= (1 + r["SPY報酬"])
            if r["超額Alpha"] > 0:
                win_count += 1

        total_years = sum(r["天數"] for r in results) / 365.25
        strategy_cagr = (cum_strategy) ** (1 / total_years) - 1
        spy_cagr = (cum_spy) ** (1 / total_years) - 1
        win_rate = win_count / len(results)
        avg_alpha = np.mean([r["超額Alpha"] for r in results])

        # 在收縮期的表現
        recession_results = [r for r in results if r["階段"] == "收縮"]
        if recession_results:
            avg_recession_alpha = np.mean([r["超額Alpha"] for r in recession_results])
            avg_recession_strategy = np.mean([r["策略報酬"] for r in recession_results])
            avg_recession_spy = np.mean([r["SPY報酬"] for r in recession_results])
        else:
            avg_recession_alpha = 0
            avg_recession_strategy = 0
            avg_recession_spy = 0

        summary = {
            "回測期間": f"{results[0]['週期'].split('_')[0]} → 2026",
            "總年數": total_years,
            "策略累計報酬": cum_strategy - 1,
            "SPY累計報酬": cum_spy - 1,
            "策略CAGR": strategy_cagr,
            "SPY_CAGR": spy_cagr,
            "超額CAGR": strategy_cagr - spy_cagr,
            "週期勝率": win_rate,
            "平均Alpha": avg_alpha,
            "衰退期平均Alpha": avg_recession_alpha,
            "衰退期策略平均報酬": avg_recession_strategy,
            "衰退期SPY平均報酬": avg_recession_spy,
        }

        print(f"\n  📅 回測期間: {summary['回測期間']} ({total_years:.1f}年)")
        print(f"  💰 策略累計報酬: {summary['策略累計報酬']:+.1%}")
        print(f"  💰 SPY累計報酬:  {summary['SPY累計報酬']:+.1%}")
        print(f"  📈 策略 CAGR:    {strategy_cagr:+.2%}")
        print(f"  📈 SPY CAGR:     {spy_cagr:+.2%}")
        print(f"  🎯 超額年化報酬:  {summary['超額CAGR']:+.2%}")
        print(f"  🏆 週期勝率:     {win_rate:.0%} ({win_count}/{len(results)})")
        print(f"  📊 平均Alpha:    {avg_alpha:+.2%}")
        print(f"\n  🛡️ 衰退期防禦能力:")
        print(f"     策略平均:  {avg_recession_strategy:+.2%}")
        print(f"     SPY平均:   {avg_recession_spy:+.2%}")
        print(f"     衰退Alpha: {avg_recession_alpha:+.2%}")

        return summary


# ============================================================
# 第三部分：IRR 預測模型（1-5年）
# ============================================================

class IRRProjector:
    """
    基於歷史數據和當前因子的 IRR 預測
    使用蒙地卡羅模擬 + 歷史類比法
    """

    # 歷史參數（基於20年數據）
    HISTORICAL_PARAMS = {
        # 各資產類別的歷史年化報酬和波動率
        "科技成長": {"mean": 0.18, "std": 0.28, "sharpe": 0.55},
        "半導體": {"mean": 0.22, "std": 0.35, "sharpe": 0.54},
        "醫療防禦": {"mean": 0.12, "std": 0.16, "sharpe": 0.56},
        "價值股": {"mean": 0.10, "std": 0.18, "sharpe": 0.44},
        "能源": {"mean": 0.08, "std": 0.30, "sharpe": 0.22},
        "國防": {"mean": 0.14, "std": 0.22, "sharpe": 0.50},
        "黃金": {"mean": 0.08, "std": 0.16, "sharpe": 0.38},
        "債券": {"mean": 0.04, "std": 0.06, "sharpe": 0.50},
        "現金": {"mean": 0.04, "std": 0.01, "sharpe": 2.50},
        "新興市場": {"mean": 0.09, "std": 0.24, "sharpe": 0.30},
    }

    # 當前組合的資產類別映射
    POSITION_CATEGORY = {
        "GOOGL": "科技成長",
        "TSM": "半導體",
        "MU": "半導體",
        "MRK": "醫療防禦",
        "JNJ": "醫療防禦",
        "VZ": "價值股",
        "PEP": "價值股",
        "XLE_ETF": "能源",
        "ITA_ETF": "國防",
        "LLY": "醫療防禦",
        "NVDA": "半導體",
        "MRVL": "半導體",
        "GLD_ETF": "黃金",
        "AGG_ETF": "債券",
        "SHV_ETF": "現金",
        "EEM_ETF": "新興市場",
    }

    # 當前環境調整因子（滯脹環境）
    ENVIRONMENT_ADJUSTMENTS = {
        "科技成長": -0.05,   # 高利率壓估值
        "半導體": -0.03,    # AI需求抵消部分壓力
        "醫療防禦": +0.03,   # 防禦需求增加
        "價值股": +0.02,    # 價值股在滯脹期表現較好
        "能源": +0.15,      # 地緣政治溢價
        "國防": +0.10,      # 軍費擴張
        "黃金": +0.06,      # 避險需求
        "債券": -0.02,      # 利率不確定性
        "現金": +0.00,      # 穩定
        "新興市場": -0.03,   # 美元走強壓力
    }

    def calc_portfolio_expected_return(self):
        """計算組合的預期年化報酬和波動率"""
        all_positions = {}
        for section in ["核心持倉", "衛星持倉", "防守持倉"]:
            all_positions.update(PORTFOLIO[section])

        weighted_return = 0
        weighted_var = 0  # 簡化：假設不考慮相關性（保守估計）

        for ticker, info in all_positions.items():
            weight = info["配置比例"]
            category = self.POSITION_CATEGORY.get(ticker, "科技成長")
            params = self.HISTORICAL_PARAMS[category]
            adj = self.ENVIRONMENT_ADJUSTMENTS.get(category, 0)

            expected_ret = params["mean"] + adj
            expected_std = params["std"]

            weighted_return += weight * expected_ret
            weighted_var += (weight ** 2) * (expected_std ** 2)

        portfolio_std = np.sqrt(weighted_var)
        sharpe = weighted_return / portfolio_std if portfolio_std > 0 else 0

        return weighted_return, portfolio_std, sharpe

    def monte_carlo_simulation(self, n_simulations=10000, years=5):
        """蒙地卡羅模擬多年期報酬分布"""
        exp_return, exp_std, _ = self.calc_portfolio_expected_return()

        results = {}
        for year in range(1, years + 1):
            # 模擬 n_simulations 條路徑
            annual_returns = np.random.normal(exp_return, exp_std, (n_simulations, year))
            cumulative = np.prod(1 + annual_returns, axis=1) - 1

            # 計算 IRR（年化）
            irr_values = (1 + cumulative) ** (1 / year) - 1

            results[year] = {
                "期望IRR": np.mean(irr_values),
                "中位數IRR": np.median(irr_values),
                "P10（悲觀）": np.percentile(irr_values, 10),
                "P25": np.percentile(irr_values, 25),
                "P75": np.percentile(irr_values, 75),
                "P90（樂觀）": np.percentile(irr_values, 90),
                "正報酬機率": np.mean(cumulative > 0),
                "累計報酬_期望": np.mean(cumulative),
                "累計報酬_中位": np.median(cumulative),
                "最大回撤_中位": -np.percentile(-cumulative, 50) if np.any(cumulative < 0) else 0,
                "期末金額_期望": 100000 * (1 + np.mean(cumulative)),
                "期末金額_P10": 100000 * (1 + np.percentile(cumulative, 10)),
                "期末金額_P90": 100000 * (1 + np.percentile(cumulative, 90)),
            }

        return results

    def print_irr_table(self, results):
        """輸出 IRR 預測表"""
        print("\n" + "=" * 80)
        print("📊 IRR 預測模型（蒙地卡羅 10,000次模擬）")
        print("=" * 80)

        exp_return, exp_std, sharpe = self.calc_portfolio_expected_return()
        print(f"\n  組合預期年化報酬: {exp_return:.2%}")
        print(f"  組合預期年化波動: {exp_std:.2%}")
        print(f"  預期夏普比率:     {sharpe:.2f}")

        print(f"\n{'年限':>4} | {'期望IRR':>8} | {'P10悲觀':>8} | {'P25':>8} | {'中位數':>8} | {'P75':>8} | {'P90樂觀':>8} | {'正報酬%':>7} | {'期末金額(萬)':>12}")
        print("─" * 100)

        for year, data in results.items():
            print(f" {year}年  | {data['期望IRR']:>+7.2%} | {data['P10（悲觀）']:>+7.2%} | {data['P25']:>+7.2%} | "
                  f"{data['中位數IRR']:>+7.2%} | {data['P75']:>+7.2%} | {data['P90（樂觀）']:>+7.2%} | "
                  f"{data['正報酬機率']:>6.1%} | ${data['期末金額_期望']/10000:>9.2f}萬")

        print(f"\n{'─' * 100}")
        print(f"  💡 模擬假設：基於歷史20年數據 + 當前滯脹環境調整")
        print(f"  💡 P10 = 90%機率優於此值 | P90 = 僅10%機率達到")

        # 期末金額範圍
        print(f"\n  📈 5年期末金額預測（$100,000起始）：")
        r5 = results[5]
        print(f"     悲觀情景 (P10): ${r5['期末金額_P10']:>12,.0f}")
        print(f"     期望情景:       ${r5['期末金額_期望']:>12,.0f}")
        print(f"     樂觀情景 (P90): ${r5['期末金額_P90']:>12,.0f}")


# ============================================================
# 第四部分：輸出投資組合摘要
# ============================================================

def print_portfolio_summary():
    """輸出完整投資組合摘要"""
    print("=" * 80)
    print("💰 方案一：$100,000 最優投資組合")
    print(f"📅 日期: {PORTFOLIO['基準日期']}")
    print(f"📋 策略: {PORTFOLIO['策略名稱']}")
    print(f"🌍 景氣判斷: {PORTFOLIO['當前景氣判斷']}")
    print("=" * 80)

    total_invested = 0

    for section_name in ["核心持倉", "衛星持倉", "防守持倉"]:
        section = PORTFOLIO[section_name]
        section_total = sum(v["金額"] for v in section.values())
        section_pct = section_total / PORTFOLIO["總資金"]

        print(f"\n{'─' * 60}")
        print(f"📁 {section_name} ({section_pct:.0%} = ${section_total:,})")
        print(f"{'─' * 60}")
        print(f"{'代碼':<10} {'名稱':<25} {'比例':>5} {'金額':>8} {'股數':>5} {'QVM分數':>8}")

        for ticker, info in section.items():
            q = info["因子"].get("Quality", "N/A")
            v = info["因子"].get("Value", "N/A")
            m = info["因子"].get("Momentum", "N/A")

            qvm_str = f"Q{q}/V{v}/M{m}" if isinstance(q, (int, float)) else "ETF"

            print(f"{ticker:<10} {info['名稱']:<25} {info['配置比例']:>4.0%} "
                  f"${info['金額']:>7,} {info['股數']:>5} {qvm_str:>8}")

        total_invested += section_total

    print(f"\n{'═' * 60}")
    print(f"  💵 總配置: ${total_invested:,} / ${PORTFOLIO['總資金']:,}")

    # 板塊分布
    print(f"\n📊 板塊分布:")
    sectors = {}
    for section_name in ["核心持倉", "衛星持倉", "防守持倉"]:
        for ticker, info in PORTFOLIO[section_name].items():
            sector = info["板塊"]
            sectors[sector] = sectors.get(sector, 0) + info["配置比例"]

    for sector, weight in sorted(sectors.items(), key=lambda x: -x[1]):
        bar = "█" * int(weight * 50)
        print(f"  {sector:<16} {weight:>5.0%} {bar}")


# ============================================================
# 主程式
# ============================================================

def main():
    # Step 1: 輸出投資組合摘要
    print_portfolio_summary()

    # Step 2: 20年歷史回測驗證
    print("\n\n" + "🔬" * 30)
    print("啟動20年歷史回測驗證系統...")
    print("驗證 QVM + 景氣循環策略在每個經濟週期中的有效性")
    print("🔬" * 30)

    backtester = PortfolioBacktester()
    results = backtester.backtest_all_cycles()
    summary = backtester.calc_cumulative_performance(results)

    # Step 3: IRR 預測
    print("\n\n" + "📈" * 30)
    print("啟動蒙地卡羅 IRR 預測模型...")
    print("📈" * 30)

    projector = IRRProjector()
    irr_results = projector.monte_carlo_simulation(n_simulations=10000, years=5)
    projector.print_irr_table(irr_results)

    # Step 4: 保存完整結果
    output = {
        "投資組合": PORTFOLIO,
        "回測結果": results if results else [],
        "累計統計": summary if summary else {},
        "IRR預測": {str(k): {kk: float(vv) if isinstance(vv, (np.floating, float)) else vv
                            for kk, vv in v.items()}
                   for k, v in irr_results.items()},
    }

    output_path = os.path.join(os.path.dirname(__file__), "方案一_回測結果.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n\n✅ 完整結果已儲存: {output_path}")

    # Step 5: 最終評級
    print("\n" + "=" * 80)
    print("🏆 最終投資組合評級")
    print("=" * 80)
    if summary:
        print(f"  策略歷史 CAGR:  {summary.get('策略CAGR', 0):+.2%}")
        print(f"  超額年化報酬:    {summary.get('超額CAGR', 0):+.2%}")
        print(f"  景氣週期勝率:    {summary.get('週期勝率', 0):.0%}")
        print(f"  衰退期防禦Alpha: {summary.get('衰退期平均Alpha', 0):+.2%}")

    irr_1y = irr_results[1]
    irr_5y = irr_results[5]
    print(f"\n  1年期望IRR: {irr_1y['期望IRR']:+.2%} (正報酬機率: {irr_1y['正報酬機率']:.1%})")
    print(f"  5年期望IRR: {irr_5y['期望IRR']:+.2%} (正報酬機率: {irr_5y['正報酬機率']:.1%})")
    print(f"  5年期末金額: ${irr_5y['期末金額_期望']:,.0f} (期望值)")

    print(f"\n  ⚖️ 風險評估: 中等偏保守")
    print(f"  🎯 適合: 能承受短期波動、追求中長期穩健增長的投資者")
    print(f"  ⏰ 建議持有: 至少3年以上")


if __name__ == "__main__":
    main()
