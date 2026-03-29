#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
方案五：對沖基金策略投資組合 — $100,000
複製頂級對沖基金策略，追求絕對報酬
多空 + 全球宏觀 + CTA趨勢 + 併購套利 + 統計套利
"""

import numpy as np
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backtest_system'))

# ============================================================
# 第一部分：投資組合定義
# ============================================================

PORTFOLIO = {
    "總資金": 100000,
    "基準日期": "2026-03-29",
    "策略名稱": "多策略對沖基金複製組合",
    "目標": "絕對報酬10-15% p.a.，最大回撤<15%，牛市跟漲熊市抗跌",

    # ===== 模塊一：核心多空股票（30%）=====
    "多空股票": {
        "配置比例": 0.30,
        "金額": 30000,
        "淨曝險": "60-80% 偏多",
        "做多": {
            "NVDA": {"金額": 5000, "理由": "5/6頂級基金重倉，AI算力壟斷", "目標": "+25%"},
            "GOOGL": {"金額": 4000, "理由": "PE 25x最低AI龍頭，RSI超賣", "目標": "+20%"},
            "META": {"金額": 3500, "理由": "PE 22x+ROE 36%，Citadel重倉", "目標": "+18%"},
            "MRK": {"金額": 3000, "理由": "Keytruda+PE 16x，防禦型成長", "目標": "+15%"},
            "XLE": {"金額": 3000, "理由": "伊朗戰爭油價溢價+YTD +40%", "目標": "+10%"},
            "ITA": {"金額": 2500, "理由": "國防預算$9,000億+，最強主題", "目標": "+15%"},
        },
        "做空": {
            "ARKK": {"金額": -3000, "理由": "高估值無利潤科技股，升息環境最脆弱", "目標": "-20%"},
            "XLRE": {"金額": -2000, "理由": "商業地產空置率創新高+高利率壓力", "目標": "-15%"},
            "XRT": {"金額": -2000, "理由": "零售業消費降級，非必需消費板塊弱勢", "目標": "-10%"},
        },
        "策略預期": {
            "做多報酬": "+15-20%",
            "做空報酬": "+5-10% (空方獲利)",
            "淨報酬": "+12-18%",
        },
    },

    # ===== 模塊二：全球宏觀（20%）=====
    "全球宏觀": {
        "配置比例": 0.20,
        "金額": 20000,
        "說明": "複製Bridgewater/Citadel的宏觀交易，多資產跨國家配置",
        "交易": {
            "做多黃金": {
                "工具": "GLD",
                "金額": 6000,
                "理由": "高盛目標$4,900/oz，央行購金+去美元化+地緣避險",
                "目標": "+20-30%",
                "來源": "Citadel 2025 Q4加倉$42億GLD",
            },
            "做多日圓": {
                "工具": "FXY (日圓ETF)",
                "金額": 3000,
                "理由": "BOJ升息至1%趨勢，日圓被嚴重低估，MUFG預測美元跌至90-95",
                "目標": "+10-15%",
            },
            "做多長期國債": {
                "工具": "TLT",
                "金額": 4000,
                "理由": "殖利率4.5%歷史高位，若衰退降息TLT將暴漲",
                "目標": "+8-15%",
            },
            "做多銅礦": {
                "工具": "COPX (銅礦ETF)",
                "金額": 3000,
                "理由": "AI資料中心+電動車+電網升級推動銅需求，供給受限",
                "目標": "+15-25%",
            },
            "做空歐洲天然氣": {
                "工具": "反向天然氣ETN",
                "金額": 2000,
                "理由": "LNG產能浪潮，TTF預計跌35%",
                "目標": "+10-20%",
            },
            "做空美元": {
                "工具": "UDN (美元空頭ETF)",
                "金額": 2000,
                "理由": "Fed降息週期+雙赤字擴大，MUFG預測DXY跌至90-95",
                "目標": "+5-10%",
            },
        },
    },

    # ===== 模塊三：CTA趨勢追蹤（15%）=====
    "CTA趨勢": {
        "配置比例": 0.15,
        "金額": 15000,
        "說明": "管理期貨策略，追蹤跨資產趨勢，危機時提供正報酬",
        "持倉": {
            "DBMF": {
                "金額": 8000,
                "名稱": "iMGP DBi Managed Futures ETF",
                "費用率": "0.85%",
                "策略": "複製前20大CTA基金的報酬",
                "歷史": "2022年+24.1%（SPY -18.1%時），2008年型危機中預期+15-25%",
                "邏輯": "最佳的危機Alpha工具，牛市平穩、熊市暴漲",
            },
            "KMLM": {
                "金額": 5000,
                "名稱": "KFA Mount Lucas Managed Futures ETF",
                "費用率": "0.92%",
                "策略": "商品+外匯+利率趨勢追蹤",
                "歷史": "2022年+30.2%（表現比DBMF更好）",
                "邏輯": "更高的商品和匯率曝險，當前環境有利",
            },
            "CTA": {
                "金額": 2000,
                "名稱": "Simplify Managed Futures Strategy ETF",
                "費用率": "0.75%",
                "策略": "系統性趨勢+carry策略",
                "邏輯": "補充分散，與DBMF/KMLM低相關",
            },
        },
        "策略預期": {
            "牛市預期": "+3-8%",
            "熊市預期": "+15-30%（危機Alpha）",
            "平均預期": "+8-12%",
        },
    },

    # ===== 模塊四：併購套利（10%）=====
    "併購套利": {
        "配置比例": 0.10,
        "金額": 10000,
        "說明": "捕捉M&A價差，與大盤相關性極低",
        "持倉": {
            "MNA": {
                "金額": 7000,
                "名稱": "IQ Merger Arbitrage ETF",
                "費用率": "0.77%",
                "策略": "系統性併購套利，持有被收購公司股票",
                "歷史CAGR": "+4-6%",
                "波動率": "3-5%",
                "邏輯": "與股市相關性<0.2，提供穩定正報酬不論牛熊",
            },
            "現金_併購機會": {
                "金額": 3000,
                "說明": "保留現金等待特殊機會（如大型破裂交易的反向投注）",
            },
        },
    },

    # ===== 模塊五：統計套利 + 動量（10%）=====
    "統計套利": {
        "配置比例": 0.10,
        "金額": 10000,
        "持倉": {
            "QAI": {
                "金額": 4000,
                "名稱": "IQ Hedge Multi-Strategy Tracker ETF",
                "費用率": "0.75%",
                "策略": "多策略對沖基金複製",
                "邏輯": "分散至多種對沖策略，降低單一策略風險",
            },
            "IWM": {
                "金額": 4000,
                "名稱": "Russell 2000 小型股ETF",
                "PE": 17.9,
                "邏輯": "Millennium大幅加倉$30億！PE 17.9x遠低於大盤25x，降息受益最大",
            },
            "配對交易_現金": {
                "金額": 2000,
                "說明": "保留現金用於配對交易機會（如XOM/CVX、GOOGL/META）",
            },
        },
    },

    # ===== 現金緩衝（5%）=====
    "現金": {
        "配置比例": 0.05,
        "金額": 5000,
        "工具": "SHV (殖利率~4%)",
    },
}


# ============================================================
# 第二部分：對沖基金策略歷史回測
# ============================================================

class HedgeFundBacktester:
    """使用HFRI指數和ETF數據回測"""

    # HFRI 各策略指數年度回報（來源: HFR官方）
    HFRI_RETURNS = {
        # year: (綜合, 多空, 宏觀, 事件驅動, 相對價值)
        2003: (0.198, 0.175, 0.180, 0.260, 0.140),
        2004: (0.090, 0.067, 0.058, 0.155, 0.075),
        2005: (0.095, 0.097, 0.090, 0.073, 0.068),
        2006: (0.128, 0.117, 0.099, 0.158, 0.124),
        2007: (0.100, 0.073, 0.116, 0.088, 0.100),
        2008: (-0.190, -0.266, -0.046, -0.217, -0.181),  # 宏觀最佳
        2009: (0.200, 0.244, 0.046, 0.260, 0.259),
        2010: (0.105, 0.092, 0.082, 0.120, 0.110),
        2011: (-0.051, -0.086, -0.031, -0.036, 0.001),
        2012: (0.063, 0.050, 0.019, 0.087, 0.090),
        2013: (0.091, 0.145, -0.001, 0.120, 0.077),
        2014: (0.031, 0.019, 0.056, 0.011, 0.039),
        2015: (-0.010, -0.003, -0.012, -0.037, 0.022),
        2016: (0.057, 0.047, 0.012, 0.109, 0.070),
        2017: (0.086, 0.113, 0.029, 0.078, 0.065),
        2018: (-0.042, -0.067, -0.035, -0.014, -0.003),
        2019: (0.102, 0.107, 0.063, 0.081, 0.085),
        2020: (0.115, 0.172, 0.058, 0.079, 0.057),
        2021: (0.104, 0.127, 0.065, 0.126, 0.048),
        2022: (-0.042, -0.065, 0.092, -0.030, 0.003),  # 宏觀+CTA大勝
        2023: (0.080, 0.095, 0.048, 0.072, 0.065),
        2024: (0.102, 0.115, 0.085, 0.090, 0.078),
        2025: (0.126, 0.135, 0.118, 0.105, 0.092),  # 16年最佳
    }

    SPY_RETURNS = {
        2003: 0.287, 2004: 0.109, 2005: 0.049, 2006: 0.158,
        2007: 0.055, 2008: -0.370, 2009: 0.265, 2010: 0.151,
        2011: 0.021, 2012: 0.160, 2013: 0.324, 2014: 0.137,
        2015: 0.014, 2016: 0.120, 2017: 0.218, 2018: -0.044,
        2019: 0.315, 2020: 0.184, 2021: 0.287, 2022: -0.181,
        2023: 0.261, 2024: 0.250, 2025: 0.120,
    }

    # CTA/管理期貨額外數據
    CTA_RETURNS = {
        2008: 0.140,   # 危機大勝
        2009: -0.012,
        2010: 0.120, 2011: -0.030, 2012: -0.020,
        2013: -0.010, 2014: 0.150, 2015: -0.010,
        2016: -0.030, 2017: 0.020, 2018: -0.060,
        2019: 0.060, 2020: 0.030, 2021: 0.100,
        2022: 0.241,  # DBMF +24.1%
        2023: -0.050, 2024: 0.080, 2025: 0.100,
    }

    def run_backtest(self):
        print("\n" + "=" * 90)
        print("📊 20年歷史回測 — 多策略對沖基金 vs SPY vs HFRI")
        print("=" * 90)

        years = sorted(self.HFRI_RETURNS.keys())
        cum_hfri = cum_spy = cum_ours = 1.0
        our_wins = 0

        results = []

        print(f"\n{'年份':>5} | {'HFRI綜合':>8} | {'SPY':>8} | {'我們的策略':>10} | {'vs SPY':>8}")
        print("─" * 60)

        for year in years:
            hfri_comp, hfri_ls, hfri_macro, hfri_event, hfri_rv = self.HFRI_RETURNS[year]
            spy_ret = self.SPY_RETURNS.get(year, 0)
            cta_ret = self.CTA_RETURNS.get(year, 0.03)

            # 我們的組合：30%多空 + 20%宏觀 + 15%CTA + 10%事件 + 10%相對價值 + 5%現金(3%) + 10% IWM
            iwm_ret = spy_ret * 1.1  # 小型股Beta約1.1
            our_ret = (0.30 * hfri_ls + 0.20 * hfri_macro + 0.15 * cta_ret +
                      0.10 * hfri_event + 0.10 * hfri_rv + 0.10 * iwm_ret * 0.5 +
                      0.05 * 0.03)

            alpha = our_ret - spy_ret
            cum_hfri *= (1 + hfri_comp)
            cum_spy *= (1 + spy_ret)
            cum_ours *= (1 + our_ret)

            if alpha > 0:
                our_wins += 1

            results.append({"year": year, "hfri": hfri_comp, "spy": spy_ret,
                           "ours": our_ret, "alpha": alpha})

            flag = "✅" if alpha > 0 else "  "
            print(f" {year} | {hfri_comp:>+7.2%} | {spy_ret:>+7.2%} | {our_ret:>+9.2%} | {alpha:>+7.2%} {flag}")

        n = len(years)
        hfri_cagr = cum_hfri ** (1/n) - 1
        spy_cagr = cum_spy ** (1/n) - 1
        ours_cagr = cum_ours ** (1/n) - 1

        ours_vol = np.std([r["ours"] for r in results])
        spy_vol = np.std([r["spy"] for r in results])
        ours_sharpe = ours_cagr / ours_vol if ours_vol > 0 else 0
        spy_sharpe = spy_cagr / spy_vol if spy_vol > 0 else 0

        ours_max_dd = min(r["ours"] for r in results)
        spy_max_dd = min(r["spy"] for r in results)

        # 危機年份
        crisis = [r for r in results if r["spy"] < -0.05]
        crisis_ours = np.mean([r["ours"] for r in crisis]) if crisis else 0
        crisis_spy = np.mean([r["spy"] for r in crisis]) if crisis else 0

        print(f"\n{'═' * 70}")
        print(f"📊 累計統計（{years[0]}-{years[-1]}，{n}年）")
        print(f"{'═' * 70}")
        print(f"  {'指標':<18} | {'我們的策略':>12} | {'HFRI綜合':>10} | {'SPY':>10}")
        print(f"  {'─' * 55}")
        print(f"  {'累計報酬':<18} | {cum_ours-1:>+11.1%} | {cum_hfri-1:>+9.1%} | {cum_spy-1:>+9.1%}")
        print(f"  {'CAGR':<18} | {ours_cagr:>+11.2%} | {hfri_cagr:>+9.2%} | {spy_cagr:>+9.2%}")
        print(f"  {'波動率':<18} | {ours_vol:>10.2%} | {'—':>10} | {spy_vol:>9.2%}")
        print(f"  {'夏普比率':<18} | {ours_sharpe:>11.2f} | {'—':>10} | {spy_sharpe:>9.2f}")
        print(f"  {'最大年虧損':<18} | {ours_max_dd:>+11.2%} | {'—':>10} | {spy_max_dd:>+9.2%}")
        print(f"  {'跑贏SPY':<18} | {our_wins:>8}/{n} | {'—':>10} | {'—':>10}")
        print(f"  {'危機年平均':<18} | {crisis_ours:>+11.2%} | {'—':>10} | {crisis_spy:>+9.2%}")

        return results, {
            "cagr": ours_cagr, "vol": ours_vol, "sharpe": ours_sharpe,
            "max_dd": ours_max_dd, "wins": our_wins, "n": n,
            "spy_cagr": spy_cagr, "spy_vol": spy_vol,
            "crisis_ours": crisis_ours, "crisis_spy": crisis_spy,
        }


def monte_carlo_irr(summary, n_sim=10000, years=5):
    """IRR預測"""
    exp_ret = summary["cagr"]
    exp_vol = summary["vol"]

    print(f"\n{'═' * 80}")
    print(f"📊 IRR 預測（蒙地卡羅 10,000次）")
    print(f"  預期年化: {exp_ret:.2%} | 波動率: {exp_vol:.2%} | 夏普: {summary['sharpe']:.2f}")
    print(f"{'═' * 80}")
    print(f"{'年限':>4} | {'期望IRR':>8} | {'P10':>8} | {'中位數':>8} | {'P90':>8} | {'正報酬%':>7} | {'期末金額':>12}")
    print("─" * 75)

    irr_results = {}
    for year in range(1, years + 1):
        annual = np.random.normal(exp_ret, exp_vol, (n_sim, year))
        # 加入尾部風險
        extreme = np.random.random((n_sim, year)) < 0.03
        annual[extreme] = np.random.normal(-0.10, 0.08, extreme.sum())

        cum = np.prod(1 + annual, axis=1) - 1
        irr = (1 + cum) ** (1/year) - 1

        d = {
            "期望IRR": float(np.mean(irr)),
            "P10": float(np.percentile(irr, 10)),
            "中位數": float(np.median(irr)),
            "P90": float(np.percentile(irr, 90)),
            "正報酬機率": float(np.mean(cum > 0)),
            "期末金額": float(100000 * (1 + np.mean(cum))),
        }
        irr_results[year] = d
        print(f" {year}年  | {d['期望IRR']:>+7.2%} | {d['P10']:>+7.2%} | {d['中位數']:>+7.2%} | "
              f"{d['P90']:>+7.2%} | {d['正報酬機率']:>6.1%} | ${d['期末金額']:>10,.0f}")

    return irr_results


def print_portfolio():
    print("=" * 80)
    print("🏦 方案五：對沖基金策略投資組合 — $100,000")
    print(f"📅 {PORTFOLIO['基準日期']} | 目標: {PORTFOLIO['目標']}")
    print("=" * 80)

    modules = [
        ("多空股票", "📈"),
        ("全球宏觀", "🌍"),
        ("CTA趨勢", "📊"),
        ("併購套利", "🔀"),
        ("統計套利", "🔬"),
    ]

    for module_name, emoji in modules:
        mod = PORTFOLIO[module_name]
        print(f"\n{'━' * 70}")
        print(f"{emoji} {module_name} ({mod['配置比例']:.0%} = ${mod['金額']:,})")
        print(f"{'━' * 70}")

        if module_name == "多空股票":
            print("  做多:")
            for t, info in mod["做多"].items():
                print(f"    {t:<8} ${info['金額']:>5,} | {info['理由']} | 目標{info['目標']}")
            print("  做空:")
            for t, info in mod["做空"].items():
                print(f"    {t:<8} ${abs(info['金額']):>5,} | {info['理由']} | 目標{info['目標']}")
        elif module_name == "全球宏觀":
            for trade_name, info in mod["交易"].items():
                print(f"  {trade_name:<12} {info['工具']:<8} ${info['金額']:>5,} | {info['理由'][:50]}")
        elif "持倉" in mod:
            for t, info in mod["持倉"].items():
                name = info.get("名稱", info.get("說明", ""))
                amt = info.get("金額", 0)
                print(f"  {t:<25} ${amt:>5,} | {name[:45]}")

    print(f"\n  現金緩衝: ${PORTFOLIO['現金']['金額']:,} ({PORTFOLIO['現金']['工具']})")


def main():
    print_portfolio()

    backtester = HedgeFundBacktester()
    results, summary = backtester.run_backtest()

    irr_results = monte_carlo_irr(summary)

    # 儲存
    output = {
        "回測摘要": {k: float(v) if isinstance(v, (np.floating, float)) else v
                    for k, v in summary.items()},
        "IRR預測": irr_results,
    }
    output_path = os.path.join(os.path.dirname(__file__), "方案五_回測結果.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n✅ 結果已儲存: {output_path}")

    # 最終評級
    print(f"\n{'═' * 70}")
    print(f"🏆 方案五最終評級")
    print(f"{'═' * 70}")
    print(f"  策略CAGR:       {summary['cagr']:+.2%}")
    print(f"  SPY CAGR:       {summary['spy_cagr']:+.2%}")
    print(f"  夏普比率:        {summary['sharpe']:.2f}")
    print(f"  最大年虧損:      {summary['max_dd']:+.2%}")
    print(f"  勝率:           {summary['wins']}/{summary['n']}")
    print(f"  危機年表現:      {summary['crisis_ours']:+.2%} (SPY: {summary['crisis_spy']:+.2%})")
    print(f"  風險等級:        中等")
    print(f"  適合:           追求絕對報酬、不想受大盤漲跌影響的投資者")


if __name__ == "__main__":
    main()
