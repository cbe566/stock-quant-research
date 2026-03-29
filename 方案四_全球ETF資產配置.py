#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
方案四：全球ETF資產配置 — $100,000 專家級配置
跨地區、跨資產、跨主題的永久型投資組合
含經典模型比較 + 20年回測 + IRR預測
"""

import numpy as np
import pandas as pd
import requests
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backtest_system'))
from config import TIINGO_API_KEY

# ============================================================
# 第一部分：投資組合定義
# ============================================================

PORTFOLIO = {
    "總資金": 100000,
    "基準日期": "2026-03-29",
    "策略名稱": "全球多資產ETF智慧配置",
    "核心理念": "地區分散+資產分散+主題衛星，參考Vanguard/BlackRock 2026年配置建議",

    # ===== 全球股票（55%）=====
    "全球股票": {
        # 美國 — 25%
        "VTV": {
            "名稱": "Vanguard 美國價值股ETF",
            "配置比例": 0.10,
            "金額": 10000,
            "費用率": "0.04%",
            "PE": 17.5,
            "殖利率": "2.4%",
            "邏輯": "Vanguard 2026首推！價值股在滯脹環境表現優於成長股，PE僅17.5x遠低於QQQ的30x",
        },
        "SCHD": {
            "名稱": "Schwab 高股息ETF",
            "配置比例": 0.08,
            "金額": 8000,
            "費用率": "0.06%",
            "殖利率": "3.4%",
            "YTD": "+11.5%",
            "邏輯": "YTD +11.5%大幅跑贏SPY(-6.76%)，品質+股息雙因子，滯脹環境最強風格",
        },
        "SMH": {
            "名稱": "VanEck 半導體ETF",
            "配置比例": 0.07,
            "金額": 7000,
            "費用率": "0.35%",
            "1年報酬": "+62.6%",
            "邏輯": "AI基礎設施核心受益，全球半導體銷售將破$1兆，長期趨勢不變",
        },

        # 歐洲 — 10%
        "VGK": {
            "名稱": "Vanguard 歐洲ETF",
            "配置比例": 0.06,
            "金額": 6000,
            "費用率": "0.09%",
            "PE": 14.2,
            "殖利率": "3.1%",
            "邏輯": "BlackRock 2026看好歐洲！PE 14x vs 美股25x，折價40%，國防+能源轉型受益",
        },
        "EZU": {
            "名稱": "iShares 歐元區ETF",
            "配置比例": 0.04,
            "金額": 4000,
            "費用率": "0.09%",
            "邏輯": "歐元區銀行股PE僅9-10x，現金流充沛，ECB暫停降息利好銀行業",
        },

        # 日本 — 6%
        "DXJ": {
            "名稱": "WisdomTree 日本匯率避險ETF",
            "配置比例": 0.06,
            "金額": 6000,
            "費用率": "0.48%",
            "1年報酬": "+45.9%",
            "邏輯": "1年暴漲45.9%！企業治理改革+BOJ升息+日圓避險=三重催化劑",
        },

        # 新興市場 — 8%
        "EEM": {
            "名稱": "iShares 新興市場ETF",
            "配置比例": 0.04,
            "金額": 4000,
            "費用率": "0.72%",
            "PE": 16.2,
            "殖利率": "2.2%",
            "邏輯": "PE 16x vs 美股25x，估值折價35%，亞洲市場持續獲得資金流入",
        },
        "INDA": {
            "名稱": "iShares 印度ETF",
            "配置比例": 0.04,
            "金額": 4000,
            "費用率": "0.65%",
            "邏輯": "印度GDP成長6.5%全球最快，人口紅利+製造業遷移+數位化轉型",
        },
    },

    # ===== 固定收益（25%）=====
    "固定收益": {
        "IEF": {
            "名稱": "iShares 7-10年美國國債ETF",
            "配置比例": 0.08,
            "金額": 8000,
            "費用率": "0.15%",
            "殖利率": "4.2%",
            "存續期": "7.3年",
            "邏輯": "Vanguard 2026最看好債券！殖利率4.2%歷史高位，降息週期啟動將有資本利得",
        },
        "LQD": {
            "名稱": "iShares 投資級公司債ETF",
            "配置比例": 0.06,
            "金額": 6000,
            "費用率": "0.14%",
            "殖利率": "5.1%",
            "邏輯": "殖利率5.1%優於國債，投資級信用品質穩定，利差尚未過度壓縮",
        },
        "TIPS": {
            "名稱": "iShares 抗通膨債券ETF",
            "配置比例": 0.05,
            "金額": 5000,
            "費用率": "0.19%",
            "邏輯": "滯脹環境必備！CPI 2.4%+關稅效應，實質殖利率仍為正",
        },
        "SHY": {
            "名稱": "iShares 1-3年美國國債ETF",
            "配置比例": 0.03,
            "金額": 3000,
            "費用率": "0.15%",
            "殖利率": "4.0%",
            "邏輯": "短期國債幾乎零風險+4%殖利率，現金替代品",
        },
        "EMB": {
            "名稱": "iShares 新興市場美元債ETF",
            "配置比例": 0.03,
            "金額": 3000,
            "費用率": "0.39%",
            "殖利率": "6.8%",
            "邏輯": "殖利率6.8%極具吸引力，分散至非美元信用市場",
        },
    },

    # ===== 另類資產（12%）=====
    "另類資產": {
        "GLDM": {
            "名稱": "SPDR 黃金迷你ETF",
            "配置比例": 0.08,
            "金額": 8000,
            "費用率": "0.10%",
            "YTD": "+47.1%",
            "邏輯": "2026年表現最強資產！地緣政治+去美元化+央行購金，高盛目標$4,900/oz",
        },
        "VNQ": {
            "名稱": "Vanguard 美國REITs ETF",
            "配置比例": 0.02,
            "金額": 2000,
            "費用率": "0.12%",
            "殖利率": "3.8%",
            "邏輯": "REITs在降息週期受益，AI資料中心REITs成長快速",
        },
        "DBC": {
            "名稱": "Invesco 大宗商品指數ETF",
            "配置比例": 0.02,
            "金額": 2000,
            "費用率": "0.87%",
            "邏輯": "原油+金屬+農產品分散曝險，伊朗戰爭推高商品價格",
        },
    },

    # ===== 主題衛星（8%）=====
    "主題衛星": {
        "ITA": {
            "名稱": "iShares 國防航太ETF",
            "配置比例": 0.04,
            "金額": 4000,
            "費用率": "0.40%",
            "1年報酬": "+38.6%",
            "邏輯": "2026年最強主題之一，全球軍費擴張趨勢持續2-3年",
        },
        "XLV": {
            "名稱": "Health Care Select ETF",
            "配置比例": 0.04,
            "金額": 4000,
            "費用率": "0.08%",
            "邏輯": "防禦型板塊+GLP-1藥物革命，滯脹環境中穩定表現",
        },
    },
}


# ============================================================
# 第二部分：經典配置模型比較
# ============================================================

CLASSIC_MODELS = {
    "60/40 傳統": {
        "配置": {"SPY": 0.60, "AGG": 0.40},
        "描述": "經典股債6:4",
    },
    "Vanguard 2026 (40/60)": {
        "配置": {"VTV": 0.20, "VEA": 0.10, "SCHD": 0.10, "AGG": 0.30, "TLT": 0.20, "TIPS": 0.10},
        "描述": "Vanguard史上首次建議40股/60債",
    },
    "全天候 (Dalio)": {
        "配置": {"SPY": 0.30, "TLT": 0.40, "IEF": 0.15, "GLD": 0.075, "DBC": 0.075},
        "描述": "Ray Dalio全天候策略",
    },
    "永久組合": {
        "配置": {"SPY": 0.25, "TLT": 0.25, "GLD": 0.25, "SHV": 0.25},
        "描述": "Harry Browne永久組合",
    },
    "我們的策略": {
        "配置": {"VTV": 0.10, "SCHD": 0.08, "SMH": 0.07, "VGK": 0.06, "EZU": 0.04,
                 "DXJ": 0.06, "EEM": 0.04, "INDA": 0.04,
                 "IEF": 0.08, "LQD": 0.06, "TIPS": 0.05, "SHY": 0.03, "EMB": 0.03,
                 "GLD": 0.08, "VNQ": 0.02, "DBC": 0.02,
                 "ITA": 0.04, "XLV": 0.04},
        "描述": "全球多資產智慧配置",
    },
}


class GlobalETFBacktester:
    """全球ETF配置回測"""

    def __init__(self):
        self.tiingo_headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Token {TIINGO_API_KEY}'
        }
        self._cache = {}

    def get_prices(self, ticker, start, end):
        cache_key = f"{ticker}_{start}_{end}"
        if cache_key in self._cache:
            return self._cache[cache_key]
        try:
            r = requests.get(
                f'https://api.tiingo.com/tiingo/daily/{ticker}/prices'
                f'?startDate={start}&endDate={end}',
                headers=self.tiingo_headers, timeout=15)
            if r.status_code == 200:
                data = r.json()
                if data:
                    df = pd.DataFrame(data)
                    df['date'] = pd.to_datetime(df['date']).dt.tz_localize(None)
                    df = df.set_index('date')
                    self._cache[cache_key] = df
                    return df
        except Exception as e:
            print(f"  ⚠ {ticker}: {e}")
        return None

    def backtest_model(self, model_name, allocation, start="2005-01-01", end="2026-03-01"):
        """回測單一配置模型"""
        annual_returns = {}

        for year in range(2005, 2026):
            y_start = f"{year}-01-01"
            y_end = f"{year}-12-31"
            port_ret = 0
            valid_weight = 0

            for ticker, weight in allocation.items():
                df = self.get_prices(ticker, y_start, y_end)
                if df is not None and len(df) > 10:
                    ret = (df['adjClose'].iloc[-1] / df['adjClose'].iloc[0]) - 1
                    port_ret += ret * weight
                    valid_weight += weight

            if valid_weight > 0.5:
                annual_returns[year] = port_ret / valid_weight

        return annual_returns

    def compare_all_models(self):
        """比較所有配置模型"""
        print("\n" + "=" * 90)
        print("📊 經典配置模型 vs 我們的策略 — 20年歷史回測")
        print("=" * 90)

        model_results = {}

        for model_name, model_info in CLASSIC_MODELS.items():
            print(f"\n  ⏳ 回測中: {model_name}...")
            returns = self.backtest_model(model_name, model_info["配置"])
            if returns:
                years = sorted(returns.keys())
                cum = 1.0
                for y in years:
                    cum *= (1 + returns[y])
                n_years = len(years)
                cagr = cum ** (1 / n_years) - 1 if n_years > 0 else 0
                vol = np.std(list(returns.values()))
                sharpe = cagr / vol if vol > 0 else 0
                max_dd = min(returns.values())
                model_results[model_name] = {
                    "returns": returns, "cagr": cagr, "vol": vol,
                    "sharpe": sharpe, "max_dd": max_dd, "cum": cum - 1,
                    "years": n_years,
                }

        # 輸出比較表
        print(f"\n{'═' * 90}")
        print(f"{'模型':<25} | {'CAGR':>7} | {'波動率':>7} | {'夏普':>6} | {'最大年虧':>8} | {'累計':>8}")
        print(f"{'─' * 90}")

        for name, data in sorted(model_results.items(), key=lambda x: -x[1]["sharpe"]):
            print(f"{name:<25} | {data['cagr']:>+6.2%} | {data['vol']:>6.2%} | "
                  f"{data['sharpe']:>5.2f} | {data['max_dd']:>+7.2%} | {data['cum']:>+7.1%}")

        return model_results


# ============================================================
# 第三部分：IRR 預測
# ============================================================

class GlobalETFIRRProjector:

    # 各資產類別歷史參數
    ASSET_PARAMS = {
        "美國價值股": {"mean": 0.10, "std": 0.16},
        "美國高股息": {"mean": 0.09, "std": 0.14},
        "半導體": {"mean": 0.18, "std": 0.30},
        "歐洲股票": {"mean": 0.07, "std": 0.18},
        "日本股票": {"mean": 0.08, "std": 0.20},
        "新興市場": {"mean": 0.08, "std": 0.22},
        "印度": {"mean": 0.10, "std": 0.24},
        "中期國債": {"mean": 0.04, "std": 0.06},
        "投資級債": {"mean": 0.05, "std": 0.07},
        "抗通膨債": {"mean": 0.03, "std": 0.05},
        "短期國債": {"mean": 0.03, "std": 0.02},
        "新興市場債": {"mean": 0.06, "std": 0.10},
        "黃金": {"mean": 0.08, "std": 0.16},
        "REITs": {"mean": 0.08, "std": 0.20},
        "大宗商品": {"mean": 0.05, "std": 0.22},
        "國防": {"mean": 0.12, "std": 0.20},
        "醫療": {"mean": 0.10, "std": 0.15},
    }

    POSITION_MAP = {
        "VTV": "美國價值股", "SCHD": "美國高股息", "SMH": "半導體",
        "VGK": "歐洲股票", "EZU": "歐洲股票", "DXJ": "日本股票",
        "EEM": "新興市場", "INDA": "印度",
        "IEF": "中期國債", "LQD": "投資級債", "TIPS": "抗通膨債",
        "SHY": "短期國債", "EMB": "新興市場債",
        "GLDM": "黃金", "VNQ": "REITs", "DBC": "大宗商品",
        "ITA": "國防", "XLV": "醫療",
    }

    def calc_portfolio_params(self):
        weighted_ret = 0
        weighted_var = 0
        all_positions = {}
        for section in ["全球股票", "固定收益", "另類資產", "主題衛星"]:
            all_positions.update(PORTFOLIO[section])

        for ticker, info in all_positions.items():
            weight = info["配置比例"]
            category = self.POSITION_MAP.get(ticker, "美國價值股")
            params = self.ASSET_PARAMS[category]
            weighted_ret += weight * params["mean"]
            weighted_var += (weight ** 2) * (params["std"] ** 2)

        return weighted_ret, np.sqrt(weighted_var)

    def monte_carlo(self, n_sim=10000, years=5):
        exp_ret, exp_vol = self.calc_portfolio_params()
        results = {}
        for year in range(1, years + 1):
            annual = np.random.normal(exp_ret, exp_vol, (n_sim, year))
            cumulative = np.prod(1 + annual, axis=1) - 1
            irr = (1 + cumulative) ** (1 / year) - 1
            results[year] = {
                "期望IRR": float(np.mean(irr)),
                "P10": float(np.percentile(irr, 10)),
                "中位數": float(np.median(irr)),
                "P90": float(np.percentile(irr, 90)),
                "正報酬機率": float(np.mean(cumulative > 0)),
                "期末金額_期望": float(100000 * (1 + np.mean(cumulative))),
            }
        return results, exp_ret, exp_vol

    def print_results(self, results, exp_ret, exp_vol):
        print(f"\n{'═' * 80}")
        print(f"📊 IRR 預測（蒙地卡羅 10,000次）")
        print(f"  組合預期年化: {exp_ret:.2%} | 波動率: {exp_vol:.2%} | 夏普: {exp_ret/exp_vol:.2f}")
        print(f"{'═' * 80}")
        print(f"{'年限':>4} | {'期望IRR':>8} | {'P10悲觀':>8} | {'中位數':>8} | {'P90樂觀':>8} | {'正報酬%':>7} | {'期末金額':>12}")
        print("─" * 75)
        for year, d in results.items():
            print(f" {year}年  | {d['期望IRR']:>+7.2%} | {d['P10']:>+7.2%} | {d['中位數']:>+7.2%} | "
                  f"{d['P90']:>+7.2%} | {d['正報酬機率']:>6.1%} | ${d['期末金額_期望']:>10,.0f}")


# ============================================================
# 主程式
# ============================================================

def print_portfolio():
    print("=" * 80)
    print("🌍 方案四：全球ETF資產配置 — $100,000")
    print(f"📅 {PORTFOLIO['基準日期']} | {PORTFOLIO['核心理念']}")
    print("=" * 80)

    for section in ["全球股票", "固定收益", "另類資產", "主題衛星"]:
        positions = PORTFOLIO[section]
        total = sum(p["金額"] for p in positions.values())
        pct = total / PORTFOLIO["總資金"]
        print(f"\n{'━' * 70}")
        print(f"📁 {section} ({pct:.0%} = ${total:,})")
        print(f"{'━' * 70}")
        for ticker, info in positions.items():
            fee = info.get('費用率', 'N/A')
            yld = info.get('殖利率', 'N/A')
            print(f"  {ticker:<8} {info['名稱']:<35} {info['配置比例']:>4.0%} ${info['金額']:>6,} 費{fee} 息{yld}")


def main():
    print_portfolio()

    # 經典模型比較回測
    backtester = GlobalETFBacktester()
    model_results = backtester.compare_all_models()

    # IRR預測
    projector = GlobalETFIRRProjector()
    irr_results, exp_ret, exp_vol = projector.monte_carlo()
    projector.print_results(irr_results, exp_ret, exp_vol)

    # 儲存
    output = {
        "投資組合": {k: v for k, v in PORTFOLIO.items() if isinstance(v, (str, int, float))},
        "模型比較": {k: {kk: float(vv) if isinstance(vv, (np.floating, float)) else vv
                       for kk, vv in v.items() if kk != "returns"}
                   for k, v in model_results.items()},
        "IRR預測": irr_results,
    }
    output_path = os.path.join(os.path.dirname(__file__), "方案四_回測結果.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n✅ 結果已儲存: {output_path}")


if __name__ == "__main__":
    main()
