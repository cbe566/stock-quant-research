"""
全球指數成分股動態抓取
======================
動態獲取四大市場的指數成分股清單，不再寫死。

美股：道瓊30 + S&P 500 + NASDAQ 100
港股：恒生指數 + 恒生中國企業指數（國企指數）
台股：台灣加權指數主要成分股
日股：日經225 + 東證主要成分股
"""

import requests
import pandas as pd
import time
import os
import ssl
import json
from io import StringIO
from datetime import datetime
from pathlib import Path

# 繞過本地 SSL 問題
ssl._create_default_https_context = ssl._create_unverified_context

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"


def _fetch_wiki_tables(url):
    """用 requests 帶 headers 抓 Wikipedia 表格，避免 403"""
    resp = requests.get(url, headers={"User-Agent": UA}, timeout=15, verify=False)
    resp.raise_for_status()
    return pd.read_html(StringIO(resp.text))

# 快取路徑（每天只抓一次）
CACHE_DIR = Path(__file__).parent / "cache"
CACHE_DIR.mkdir(exist_ok=True)


def _today_str():
    return datetime.now().strftime("%Y%m%d")


def _load_cache(name):
    """讀取當天快取"""
    cache_file = CACHE_DIR / f"{name}_{_today_str()}.json"
    if cache_file.exists():
        with open(cache_file, "r") as f:
            return json.load(f)
    return None


def _save_cache(name, data):
    """儲存當天快取"""
    cache_file = CACHE_DIR / f"{name}_{_today_str()}.json"
    with open(cache_file, "w") as f:
        json.dump(data, f)


# ==================== 美股 ====================

def get_sp500():
    """S&P 500 成分股（從 Wikipedia 抓取）"""
    cached = _load_cache("sp500")
    if cached:
        return cached
    try:
        url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
        tables = _fetch_wiki_tables(url)
        df = tables[0]
        tickers = df["Symbol"].str.replace(".", "-", regex=False).tolist()
        _save_cache("sp500", tickers)
        print(f"  S&P 500: {len(tickers)} 隻")
        return tickers
    except Exception as e:
        print(f"  S&P 500 抓取失敗: {e}")
        return []


def get_nasdaq100():
    """NASDAQ 100 成分股"""
    cached = _load_cache("nasdaq100")
    if cached:
        return cached
    try:
        url = "https://en.wikipedia.org/wiki/Nasdaq-100"
        tables = _fetch_wiki_tables(url)
        for t in tables:
            # 找有 Ticker 欄的表格且行數 >= 90
            str_cols = [str(c) for c in t.columns]
            for col in str_cols:
                if "Ticker" in col or "ticker" in col:
                    tickers = t[col].dropna().astype(str).str.replace(".", "-", regex=False).tolist()
                    tickers = [t for t in tickers if t.isalpha() or "-" in t]
                    if len(tickers) >= 90:
                        _save_cache("nasdaq100", tickers)
                        print(f"  NASDAQ 100: {len(tickers)} 隻")
                        return tickers
        return []
    except Exception as e:
        print(f"  NASDAQ 100 抓取失敗: {e}")
        return []


def get_djia():
    """道瓊工業平均指數 30 成分股"""
    cached = _load_cache("djia")
    if cached:
        return cached
    try:
        url = "https://en.wikipedia.org/wiki/Dow_Jones_Industrial_Average"
        tables = _fetch_wiki_tables(url)
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            if "symbol" in cols:
                col = [c for c in t.columns if "symbol" in str(c).lower()][0]
                tickers = t[col].dropna().tolist()
                tickers = [str(t).strip() for t in tickers if str(t).strip().isalpha()]
                if len(tickers) >= 25:
                    _save_cache("djia", tickers)
                    print(f"  道瓊30: {len(tickers)} 隻")
                    return tickers
        return []
    except Exception as e:
        print(f"  道瓊30 抓取失敗: {e}")
        return []


def get_us_stocks():
    """美股全部：道瓊 + S&P 500 + NASDAQ 100（去重）"""
    print("抓取美股成分股...")
    all_tickers = set()
    all_tickers.update(get_djia())
    time.sleep(0.5)
    all_tickers.update(get_sp500())
    time.sleep(0.5)
    all_tickers.update(get_nasdaq100())
    result = sorted(list(all_tickers))
    print(f"  美股合計（去重）: {len(result)} 隻")
    return result


# ==================== 港股 ====================

def get_hsi():
    """恒生指數成分股"""
    cached = _load_cache("hsi")
    if cached:
        return cached
    try:
        url = "https://en.wikipedia.org/wiki/Hang_Seng_Index"
        tables = _fetch_wiki_tables(url)
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            if any("ticker" in c or "stock code" in c or "code" in c for c in cols):
                code_col = [c for c in t.columns if any(k in str(c).lower() for k in ["ticker", "stock code", "code"])][0]
                codes = t[code_col].dropna().tolist()
                tickers = []
                for c in codes:
                    c_str = str(c).strip().replace("SEHK:\xa0", "").replace("SEHK: ", "").strip()
                    try:
                        num = int(c_str)
                        tickers.append(f"{num:04d}.HK")
                    except ValueError:
                        continue
                if len(tickers) >= 20:
                    _save_cache("hsi", tickers)
                    print(f"  恒生指數: {len(tickers)} 隻")
                    return tickers
        return []
    except Exception as e:
        print(f"  恒生指數抓取失敗: {e}")
        return []


def get_hscei():
    """恒生中國企業指數（國企指數）成分股"""
    cached = _load_cache("hscei")
    if cached:
        return cached
    try:
        url = "https://en.wikipedia.org/wiki/Hang_Seng_China_Enterprises_Index"
        tables = _fetch_wiki_tables(url)
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            if any("ticker" in c or "stock code" in c or "code" in c for c in cols):
                code_col = [c for c in t.columns if any(k in str(c).lower() for k in ["ticker", "stock code", "code"])][0]
                codes = t[code_col].dropna().tolist()
                tickers = []
                for c in codes:
                    c_str = str(c).strip().replace("SEHK:\xa0", "").replace("SEHK: ", "").strip()
                    try:
                        num = int(c_str)
                        tickers.append(f"{num:04d}.HK")
                    except ValueError:
                        continue
                if len(tickers) >= 10:
                    _save_cache("hscei", tickers)
                    print(f"  國企指數: {len(tickers)} 隻")
                    return tickers
        return []
    except Exception as e:
        print(f"  國企指數抓取失敗: {e}")
        return []


def get_hk_stocks():
    """港股全部：恒指 + 國指（去重）"""
    print("抓取港股成分股...")
    all_tickers = set()
    all_tickers.update(get_hsi())
    time.sleep(0.5)
    all_tickers.update(get_hscei())
    result = sorted(list(all_tickers))
    print(f"  港股合計（去重）: {len(result)} 隻")
    return result


# ==================== 台股 ====================

def get_twse():
    """台灣加權指數全部上市股票（透過 TWSE Open API 動態抓取）"""
    cached = _load_cache("twse_full")
    if cached:
        print(f"  台股（快取）: {len(cached)} 隻")
        return cached

    try:
        url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
        resp = requests.get(url, timeout=15, verify=False)
        if resp.status_code == 200:
            data = resp.json()
            tickers = set()
            for item in data:
                code = item.get("Code", "")
                # 只取4位數字代碼（排除ETF等）
                if code.isdigit() and len(code) == 4:
                    tickers.add(f"{code}.TW")

            result = sorted(list(tickers))
            _save_cache("twse_full", result)
            print(f"  台股加權指數: {len(result)} 隻")
            return result
    except Exception as e:
        print(f"  台股 TWSE API 失敗: {e}")

    # Fallback：使用靜態清單
    fallback = ["2330.TW", "2454.TW", "2317.TW", "2303.TW", "2881.TW", "2882.TW"]
    print(f"  台股（備援）: {len(fallback)} 隻")
    return fallback


def get_tw_stocks():
    """台股全部"""
    print("抓取台股成分股...")
    return get_twse()


# ==================== 日股 ====================

def get_jpx_prime():
    """從 JPX 官方抓取東證 Prime + Standard 全部上市股"""
    cached = _load_cache("jpx_full")
    if cached:
        print(f"  日股（快取）: {len(cached)} 隻")
        return cached

    try:
        url = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"
        resp = requests.get(url, headers={"User-Agent": UA}, timeout=15, verify=False)
        if resp.status_code == 200:
            import pandas as _pd
            from io import BytesIO
            df = _pd.read_excel(BytesIO(resp.content))
            market_col = "市場・商品区分"
            code_col = "コード"

            # 只取 Prime + Standard（排除 ETF、REIT、Growth）
            prime_std = df[df[market_col].str.contains("プライム|スタンダード", na=False)]
            tickers = []
            for c in prime_std[code_col].dropna().tolist():
                try:
                    num = int(c)
                    if 1000 <= num <= 9999:
                        tickers.append(f"{num}.T")
                except (ValueError, TypeError):
                    continue

            result = sorted(list(set(tickers)))
            _save_cache("jpx_full", result)
            print(f"  日股 Prime+Standard: {len(result)} 隻")
            return result
    except Exception as e:
        print(f"  JPX API 失敗: {e}")

    # Fallback：靜態日經225
    return _get_nikkei225_static()


def _get_nikkei225_static():
    """日經225 靜態清單（備援）"""
    tickers = [
        # 科技 / 精密
        "6758.T", "6861.T", "6857.T", "7741.T", "6723.T",
        "6981.T", "7735.T", "8035.T", "6146.T", "6976.T",
        "7751.T", "7752.T", "6645.T", "6702.T", "6753.T",
        "9766.T", "4689.T", "4755.T", "7731.T", "7733.T",
        # 汽車 / 運輸機械
        "7203.T", "7267.T", "7269.T", "7201.T", "7211.T",
        "6902.T", "7261.T", "7270.T", "7272.T", "4543.T",
        # 金融
        "8306.T", "8316.T", "8411.T", "8766.T", "8750.T",
        "8604.T", "8601.T", "8630.T", "8725.T", "8795.T",
        # 消費 / 零售
        "9983.T", "9984.T", "4452.T", "2802.T", "7974.T",
        "2914.T", "2801.T", "2871.T", "2502.T", "2503.T",
        # 電子 / 電機
        "6501.T", "6503.T", "6504.T", "6752.T", "6762.T",
        "6770.T", "6479.T", "6954.T", "6971.T", "6506.T",
        # 化學
        "4063.T", "4005.T", "4021.T", "4042.T", "4043.T",
        "4061.T", "4183.T", "4188.T", "4208.T", "4631.T",
        # 醫藥
        "4502.T", "4503.T", "4568.T", "4519.T", "4578.T",
        "4507.T", "4506.T", "4151.T", "4523.T", "4530.T",
        # 機械 / 工業
        "7011.T", "6301.T", "6367.T", "6302.T", "6305.T",
        "6361.T", "6471.T", "7013.T", "6113.T", "6103.T",
        # 通訊 / 公用事業
        "9432.T", "9433.T", "9434.T", "9613.T", "9501.T",
        "9502.T", "9503.T", "9531.T", "9532.T", "4324.T",
        # 商社 / 貿易
        "8001.T", "8058.T", "8031.T", "8002.T", "8015.T",
        "8053.T", "8308.T", "8309.T", "8354.T", "8355.T",
        # 鋼鐵 / 非鐵金屬
        "5401.T", "5411.T", "5406.T", "5713.T", "5711.T",
        "5714.T", "5801.T", "5802.T", "5803.T", "3436.T",
        # 建設 / 不動產
        "1801.T", "1802.T", "1803.T", "1808.T", "1812.T",
        "1721.T", "1928.T", "1925.T", "8801.T", "8802.T",
        "8830.T", "3289.T",
        # 運輸 / 物流
        "9020.T", "9021.T", "9022.T", "9005.T", "9007.T",
        "9008.T", "9009.T", "9064.T", "9101.T", "9104.T",
        "9107.T", "9147.T",
        # 紙漿 / 橡膠 / 玻璃
        "3861.T", "3863.T", "5101.T", "5108.T", "5201.T",
        "5202.T", "5214.T", "5232.T", "5233.T",
        # 食品
        "2282.T", "2269.T", "4204.T", "2531.T",
        # 紡織 / 其他製造
        "3401.T", "3402.T", "3405.T", "3101.T", "7762.T",
        "7912.T", "4901.T", "4902.T", "6594.T",
        # 服務 / 其他
        "2413.T", "4307.T", "6098.T", "6178.T", "9602.T",
        "2432.T", "4704.T", "9735.T",
    ]

    result = sorted(list(set(tickers)))
    _save_cache("nikkei225_full", result)
    print(f"  日經225+東證: {len(result)} 隻")
    return result


def get_jp_stocks():
    """日股全部：東證 Prime + Standard"""
    print("抓取日股成分股...")
    return get_jpx_prime()


# ==================== 統一入口 ====================

def get_all_constituents(cloud_mode=False):
    """
    獲取四大市場全部成分股
    cloud_mode=True 時精簡股池，確保 GitHub Actions 25 分鐘內完成
    """
    print("=" * 50)
    print("  動態抓取全球指數成分股")
    if cloud_mode:
        print("  ☁️ 雲端精簡模式：僅抓取主要指數成分股")
    print("=" * 50)

    if cloud_mode:
        # 雲端模式：精簡股池（~885 隻 → 可在 20 分鐘內完成）
        # 美股：S&P 500（~500 隻）
        # 港股：恒指+國指（~100 隻）
        # 台股：台灣50 + 中型100 精選（~150 隻）
        # 日股：日經225（~225 隻）
        result = {
            "美股": {"stocks": get_us_stocks(), "benchmark": "SPY", "currency": "USD"},
            "港股": {"stocks": get_hk_stocks(), "benchmark": "^HSI", "currency": "HKD"},
            "台股": {"stocks": _get_tw_top(), "benchmark": "^TWII", "currency": "TWD"},
            "日股": {"stocks": _get_nikkei225_static(), "benchmark": "^N225", "currency": "JPY"},
        }
    else:
        result = {
            "美股": {"stocks": get_us_stocks(), "benchmark": "SPY", "currency": "USD"},
            "港股": {"stocks": get_hk_stocks(), "benchmark": "^HSI", "currency": "HKD"},
            "台股": {"stocks": get_tw_stocks(), "benchmark": "^TWII", "currency": "TWD"},
            "日股": {"stocks": get_jp_stocks(), "benchmark": "^N225", "currency": "JPY"},
        }

    total = sum(len(v["stocks"]) for v in result.values())
    print(f"\n總計：{total} 隻股票")
    print("=" * 50)

    return result


def _get_tw_top():
    """台股精選：台灣50 + 中型100 主要成分股（雲端模式用）"""
    cached = _load_cache("tw_top")
    if cached:
        print(f"  台股精選（快取）: {len(cached)} 隻")
        return cached

    # 台灣50 + 中型100 的代表性成分股
    tw_top = [
        # 台灣50 核心
        "2330.TW", "2454.TW", "2317.TW", "2308.TW", "2303.TW",
        "2382.TW", "2891.TW", "2881.TW", "2882.TW", "2886.TW",
        "2884.TW", "2885.TW", "3711.TW", "2412.TW", "1303.TW",
        "1301.TW", "1326.TW", "2002.TW", "1101.TW", "2912.TW",
        "5871.TW", "5880.TW", "2207.TW", "3045.TW", "2603.TW",
        "2880.TW", "6505.TW", "4904.TW", "3008.TW", "2301.TW",
        "4938.TW", "2357.TW", "2395.TW", "3034.TW", "2327.TW",
        "6669.TW", "2379.TW", "5876.TW", "2345.TW", "1216.TW",
        "2892.TW", "9910.TW", "2890.TW", "3037.TW", "1590.TW",
        "2887.TW", "6446.TW", "2883.TW", "3231.TW", "2105.TW",
        # 中型100 精選（市值較大者）
        "2347.TW", "3661.TW", "2049.TW", "3443.TW", "8069.TW",
        "2377.TW", "6239.TW", "3017.TW", "2474.TW", "6285.TW",
        "2408.TW", "3529.TW", "2344.TW", "5274.TW", "1477.TW",
        "2201.TW", "3023.TW", "8046.TW", "2542.TW", "1504.TW",
        "3044.TW", "2383.TW", "6176.TW", "6770.TW", "3653.TW",
        "2618.TW", "9945.TW", "6415.TW", "1802.TW", "2101.TW",
        "5269.TW", "3702.TW", "6531.TW", "2356.TW", "4966.TW",
        "6488.TW", "8454.TW", "3035.TW", "2360.TW", "6271.TW",
        "1476.TW", "1102.TW", "3532.TW", "6409.TW", "2324.TW",
        "6789.TW", "3665.TW", "2633.TW", "4958.TW", "2615.TW",
        # 熱門個股補充
        "3706.TW", "2609.TW", "2610.TW", "1605.TW", "2376.TW",
        "2353.TW", "3533.TW", "8150.TW", "6547.TW", "3036.TW",
        "2337.TW", "6443.TW", "3714.TW", "2404.TW", "1227.TW",
        "1304.TW", "1402.TW", "2027.TW", "2014.TW", "9904.TW",
        "5534.TW", "2388.TW", "2354.TW", "2823.TW", "6472.TW",
        "3228.TW", "2492.TW", "6550.TW", "2385.TW", "2727.TW",
    ]
    result = sorted(list(set(tw_top)))
    _save_cache("tw_top", result)
    print(f"  台股精選: {len(result)} 隻")
    return result


if __name__ == "__main__":
    markets = get_all_constituents()
    for name, config in markets.items():
        print(f"\n{name}: {len(config['stocks'])} 隻")
        print(f"  前5: {config['stocks'][:5]}")
        print(f"  後5: {config['stocks'][-5:]}")
