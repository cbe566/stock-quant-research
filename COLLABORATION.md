# 跨終端協作指南

> 多個 Claude 終端同時工作時的協作規範與未來路線圖
> 最後更新：2026-03-30

---

## 一、當前系統狀態（已完成）

| 模組 | 狀態 | 負責終端 |
|:---|:---:|:---|
| 每日篩選引擎（916隻 / 8線程 / 5分鐘） | ✅ | 終端 A |
| GitHub Actions 自動化排程 | ✅ | 終端 A |
| Cloudflare D1 資料庫 | ✅ | 終端 A |
| Cloudflare Worker REST API | ✅ | 終端 A |
| upload_to_d1.py 自動上傳 | ✅ | 終端 A |
| GitHub Secrets 安全化（移除 fallback） | ✅ | 終端 B |
| 郵件模組（send_daily_email.py） | ⚠️ 已寫但未整合到 Actions | — |
| PDF 報告（generate_report_pdf.py） | ⚠️ 已寫但僅限本地（需 reportlab） | — |

---

## 二、協作規範

### Git 衝突避免
- **修改前先 `git pull --rebase`**
- **不要同時修改同一個檔案**，特別是：
  - `.github/workflows/daily_screening.yml`
  - `daily_screening.py`
  - `worker-api/src/index.js`
- 如果衝突了：`git stash → git pull --rebase → git stash pop`

### 分工原則
- **一個終端一個模組**，完成後再交給下一個
- 跨模組依賴用檔案溝通（此文件 + `ARCHITECTURE.md`）
- 部署到 Cloudflare 前先確認 Worker 版本一致：`cd worker-api && npx wrangler deploy`

### Secrets / Keys 管理
- 所有 API Keys 已存入 GitHub Secrets（7 個）
- Cloudflare Worker API Key: `stock-screening-2026`
- **不要把 Keys 寫死在程式碼裡**，一律從環境變數讀取

---

## 三、未來路線圖（優先順序）

### P0 — 核心流程完善（本週）

#### 1. 郵件自動發送
- **現狀**: `send_daily_email.py` 已寫好 HTML 模板，但沒接入 Actions
- **做法**: 在 Actions 加一步，用 Gmail API 或 MCP Gmail 工具寄送
- **注意**:
  - 需要 Gmail OAuth 或 App Password
  - 郵件正文 = TOP 3 摘要，附件 = PDF 完整報告
  - 股票代碼要防超連結（零寬空格處理）

#### 2. PDF 報告雲端生成
- **現狀**: `generate_report_pdf.py` 用 reportlab，但 GitHub Actions 沒裝中文字體
- **做法**:
  - Actions 安裝 `reportlab` + 下載中文字體（Noto Sans CJK）
  - 或改用 Worker 端生成（但 Workers 不支援 Python）
- **建議**: Actions 安裝字體最簡單

#### 3. 篩選報告品質提升
- **現狀**: 週末跑出的數據是上週五的收盤價，不需要每天跑
- **做法**: 加入交易日判斷，週末跳過或只用快取數據
- **注意**: 美股/港股/台股/日股休市日不同

### P1 — 數據深化（本月）

#### 4. 歷史回填
- 目前 D1 只有 3/29 和 3/30 的數據
- 可以本地跑歷史篩選，然後批次上傳舊數據
- 有歷史數據後 `/api/history/:ticker` 和 `/api/summary` 才有意義

#### 5. 個股深度分析 API
- 新增 `/api/detail/:ticker` 端點
- 返回：完整因子分數、技術指標快照、分析師目標價、F-Score 明細
- 數據來自 `screening_results` 表 + 額外欄位

#### 6. 告警機制
- 當某隻股票得分劇烈變化（例如從 +5 跌到 -3），自動通知
- 可用 Worker Cron Trigger 每天比較昨今數據
- 通知管道：Email / Telegram Bot

### P2 — 前端介面（下月）

#### 7. 前端 Dashboard
- 用 Cloudflare Pages 部署靜態前端
- 技術選型：React / Vue + TailwindCSS
- 功能：
  - 四大市場篩選結果瀏覽
  - 個股得分走勢圖
  - 市場熱力圖
  - 搜尋與篩選過濾器
- API 已就緒，前端直接呼叫 Worker API

#### 8. 回測系統整合
- `backtest_system/` 已有回測框架
- 把回測結果也存入 D1，做歷史績效追蹤
- 驗證篩選模型的實際預測準確率

### P3 — 擴展（未來）

#### 9. 更多市場
- 新加坡 STI、印度 NIFTY、韓國 KOSPI
- 雲端模式下每增加一個市場 ~1-2 分鐘

#### 10. AI 加持
- 用 Workers AI 對篩選結果做自然語言摘要
- 每隻推薦股票自動生成一段分析理由
- 接入 Claude API 做深度研報

---

## 四、檔案職責對照表

| 檔案 | 職責 | 可以動嗎 |
|:---|:---|:---:|
| `daily_screening.py` | 主篩選腳本 | ⚠️ 核心檔案，修改要小心 |
| `cloud_data_engine.py` | 多源數據引擎 | ⚠️ 線程安全邏輯，別動鎖 |
| `index_constituents.py` | 成分股抓取 | ✅ 可加新市場 |
| `screening_engine.py` | 舊版多線程引擎（本地用） | ✅ 不影響雲端 |
| `upload_to_d1.py` | 上傳到 D1 | ✅ |
| `generate_report_pdf.py` | PDF 生成 | ✅ |
| `send_daily_email.py` | 郵件發送 | ✅ |
| `worker-api/src/index.js` | Worker API | ⚠️ 改完要 `npx wrangler deploy` |
| `.github/workflows/daily_screening.yml` | Actions 排程 | ⚠️ 改完要測試 |
| `ARCHITECTURE.md` | 系統架構文件 | ✅ 隨時更新 |
| `COLLABORATION.md` | 此檔案 | ✅ 隨時更新 |

---

## 五、快速指令參考

```bash
# 部署 Worker
cd worker-api && npx wrangler deploy

# 手動觸發篩選
gh workflow run "每日全球股票篩選" --ref main

# 查看最新 workflow
gh run list --limit 3

# 手動上傳數據到 D1
python3 upload_to_d1.py

# 查看 D1 數據
# 透過 API: curl https://stock-screening-api.stock-quant.workers.dev/api/latest

# Git 安全推送
git pull --rebase && git push
```
