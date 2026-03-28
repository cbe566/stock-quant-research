"""
全方位量化回測系統 — 配置檔
"""

# ==================== API Keys ====================
FINNHUB_API_KEY = "d6nsik9r01qse5qmtl4gd6nsik9r01qse5qmtl50"
FRED_API_KEY = "64b05f3e52bd84f8fb9f6f6e375c31d1"
ALPHA_VANTAGE_API_KEY = "CYSEDIRPDGYNRD3H"
FMP_API_KEY = "GEouSGBbAoOgnERMR0GjENhKwFkxeEeW"
SIMFIN_API_KEY = "e152f2dd-db67-4f4b-bfd9-2da7742908e0"
TIINGO_API_KEY = "981e9f863c8997d8e4ac01ca269c033295e4c1bb"
NEWSAPI_KEY = "pub_29a35def043f4becb3d8f65230defe8b"
POLYGON_API_KEY = "a1AkCn5B53WwqFfUeI2c5D5_NkY2BY49"
TWELVE_DATA_API_KEY = "685ced6c2b694a12b359beeec084e9cf"
EOD_API_KEY = "69c7f5b716ce48.31162491"
MARKETAUX_API_KEY = "V8xTWZIiaj2dtxt2ULUyS6Z7gQKARSygVRoYXQg6"
OPENEXCHANGE_API_KEY = "78fd1cd6d6b449e5a7861296b72ebb50"
TUSHARE_TOKEN = "cbc8a7607966ff51f3e7f1ee40df823aa1b8ab7a5f353412841b7a9d"

# 回測時間範圍
BACKTEST_START = "2020-01-01"
BACKTEST_END = "2025-12-31"

# 測試股票池
US_STOCKS = [
    # 科技巨頭
    "AAPL", "MSFT", "GOOGL", "AMZN", "NVDA", "META", "TSLA",
    # 金融
    "JPM", "BAC", "GS", "MS", "BRK-B",
    # 醫療
    "JNJ", "UNH", "PFE", "ABBV", "LLY",
    # 消費
    "WMT", "PG", "KO", "PEP", "MCD",
    # 工業/能源
    "XOM", "CVX", "CAT", "BA", "GE",
    # 半導體
    "AMD", "INTC", "AVGO", "QCOM", "MU",
    # 高成長
    "CRM", "NFLX", "ADBE", "NOW", "PANW",
]

HK_STOCKS = [
    # 恒指成分股
    "0700.HK", "9988.HK", "1810.HK", "0005.HK", "0941.HK",
    "2318.HK", "0388.HK", "3690.HK", "9618.HK", "1024.HK",
    "0027.HK", "0883.HK", "2020.HK", "0175.HK", "1211.HK",
]

# 指數基準
BENCHMARKS = {
    "US": "SPY",
    "HK": "^HSI",
}

# 策略參數
STRATEGY_PARAMS = {
    # 多因子選股
    "multifactor": {
        "quality_roe_min": 0.10,        # ROE > 10%
        "quality_debt_max": 0.70,       # 負債率 < 70%
        "value_pe_max": 25,             # P/E < 25
        "value_pb_max": 5,              # P/B < 5
        "momentum_lookback": 126,       # 6個月動量回看期
        "momentum_exclude_recent": 21,  # 排除最近1個月（避免短期反轉）
        "rebalance_days": 63,           # 每季調倉
    },
    # 技術面策略
    "technical": {
        "ma_short": 20,
        "ma_long": 50,
        "ma_trend": 200,
        "rsi_period": 14,
        "rsi_oversold": 30,
        "rsi_overbought": 70,
        "macd_fast": 12,
        "macd_slow": 26,
        "macd_signal": 9,
        "bb_period": 20,
        "bb_std": 2,
        "atr_period": 14,
    },
    # 動量策略
    "momentum": {
        "lookback_months": [1, 3, 6, 12],
        "holding_period": 21,  # 持有1個月
        "top_pct": 0.2,        # 買入前20%
        "bottom_pct": 0.2,     # 做空後20%（如果可以）
    },
    # 均值回歸
    "mean_reversion": {
        "lookback": 20,
        "entry_zscore": -2.0,
        "exit_zscore": 0.0,
        "stop_zscore": -3.5,
    },
}

# 風控參數
RISK_PARAMS = {
    "max_position_pct": 0.10,     # 單一持倉最大10%
    "max_drawdown_stop": 0.15,    # 組合回撤15%停止
    "single_stop_loss": 0.08,     # 個股止損8%
    "trailing_stop": 0.12,        # 追蹤止損12%
    "max_positions": 20,          # 最大持倉數量
    "transaction_cost": 0.001,    # 交易成本0.1%
}
