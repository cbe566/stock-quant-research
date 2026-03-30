# 系統架構文件 — 股票量化篩選系統

> 供跨終端協作參考，請勿刪除

## 架構總覽

```
┌─────────────────────────────────────────────────┐
│              GitHub Actions (每日06:28)           │
│                                                   │
│  1. daily_screening.py — 篩選916隻股票（~5min）    │
│  2. 生成 Markdown + JSON 報告                     │
│  3. upload_to_d1.py — 推送結果到 Cloudflare D1    │
│  4. git commit + push 報告到 GitHub               │
└──────────────┬───────────────────┬────────────────┘
               │                   │
               ▼                   ▼
┌──────────────────────┐  ┌─────────────────────────┐
│   GitHub Repository   │  │  Cloudflare Worker API   │
│   (報告存檔 + 程式碼)  │  │  stock-screening-api     │
└──────────────────────┘  │                           │
                          │  GET /api/latest          │
                          │  GET /api/top/:market     │
                          │  GET /api/stock/:ticker   │
                          │  GET /api/summary         │
                          │  GET /api/history/:ticker │
                          │  POST /api/upload         │
                          └─────────┬─────────────────┘
                                    │
                                    ▼
                          ┌─────────────────────────┐
                          │   Cloudflare D1 (APAC)   │
                          │   stock-screening-db     │
                          │                           │
                          │   screening_results 表    │
                          │   market_summary 表       │
                          └─────────────────────────┘
```

## 關鍵元件

### 1. 每日篩選引擎
- **檔案**: `daily_screening.py`
- **模式**: 雲端（GitHub Actions）/ 本地 雙模式
- **雲端股池**: 916 隻（美股517 + 港股95 + 台股130 + 日股174）
- **並行**: ThreadPoolExecutor 8 線程
- **時間預算**: 22 分鐘（留 buffer 給報告+上傳+push）
- **依賴**: `index_constituents.py`（成分股抓取）, `cloud_data_engine.py`（多源數據引擎）

### 2. 雲端數據引擎
- **檔案**: `cloud_data_engine.py`
- **5 層備援**: yfinance → Yahoo v8 API → 交易所 API → FMP/Tiingo/TwelveData → Stooq
- **線程安全**: `threading.Lock` 保護快取, `threading.local()` 獨立 session
- **快取**: price_cache + info_cache（同一股票不重複抓取）

### 3. Cloudflare Worker API
- **原始碼**: `worker-api/src/index.js`
- **URL**: https://stock-screening-api.stock-quant.workers.dev
- **D1 綁定**: `env.DB` → stock-screening-db
- **API Key**: `env.API_KEY`（透過 wrangler.jsonc vars 或 Secrets 設定）
- **部署**: `cd worker-api && npx wrangler deploy`

### 4. D1 資料庫
- **ID**: `a62e5ed5-0705-4c7c-a5f5-b4e0e8957cee`
- **區域**: APAC
- **表結構**:
  - `screening_results`: 每隻股票的篩選評分（date + ticker 唯一鍵）
  - `market_summary`: 每日市場概覽統計
- **索引**: date, market, ticker, (date + total_score DESC)

### 5. 上傳腳本
- **檔案**: `upload_to_d1.py`
- **觸發**: GitHub Actions 篩選完成後自動執行
- **方式**: POST JSON 到 Worker API `/api/upload`
- **認證**: Bearer token（WORKER_API_KEY）

## GitHub Secrets 需求

以下 Secrets 必須在 GitHub repo 設定（Settings → Secrets → Actions）：

| Secret 名稱 | 用途 | 當前值 |
|:---|:---|:---|
| `WORKER_API_KEY` | D1 上傳認證 | `stock-screening-2026` |
| `FMP_API_KEY` | Financial Modeling Prep API | `GEouSGBbAoOgnERMR0GjENhKwFkxeEeW` |
| `ALPHA_VANTAGE_API_KEY` | Alpha Vantage API | `CYSEDIRPDGYNRD3H` |
| `TWELVE_DATA_API_KEY` | Twelve Data API | `685ced6c2b694a12b359beeec084e9cf` |
| `FINNHUB_API_KEY` | Finnhub API | `d6nsik9r01qse5qmtl4gd6nsik9r01qse5qmtl50` |
| `TIINGO_API_KEY` | Tiingo API | `981e9f863c8997d8e4ac01ca269c033295e4c1bb` |
| `EOD_API_KEY` | EOD Historical Data | `69c7f5b716ce48.31162491` |

**重要**: 如果 workflow 中移除了 fallback 預設值（`|| 'xxx'`），則這些 Secrets **必須**先在 GitHub 設定好，否則篩選會因為 API Key 為空而失敗。

## Cloudflare 帳號資訊

- **帳號 ID**: `2d22807a3d36b4e465d9313b69c4ae77`
- **Email**: Backup901012@gmail.com
- **Workers 子網域**: `stock-quant.workers.dev`
- **方案**: 免費（$0/月）

## 注意事項

1. **不要同時修改 workflow**: 兩個終端同時 push 會衝突，用 `git pull --rebase` 解決
2. **D1 SQL 變數上限**: 每條 SQL 最多 100 個參數，用 `env.DB.batch()` 批次處理
3. **workers.dev SSL**: 新建子網域需要 5-10 分鐘 SSL 才生效
4. **台股/日股數據來源**: 台股靠 TWSE API，日股靠 JPX 官方 XLS（需要 openpyxl）
5. **美股/港股成分股**: 從 Wikipedia 表格抓取（需要 lxml + html5lib）
