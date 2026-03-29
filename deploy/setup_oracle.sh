#!/bin/bash
# ================================================
# Oracle Cloud VPS 一鍵部署腳本
# 每日全球股票篩選系統（4,814 隻股票）
# ================================================
# 用法：ssh 到 VPS 後執行
#   curl -sSL <this_script_url> | bash
# 或：
#   bash setup_oracle.sh
# ================================================

set -e

echo "================================================"
echo "  每日全球股票篩選系統 — Oracle Cloud 部署"
echo "================================================"

# 1. 系統更新
echo "[1/7] 系統更新..."
sudo apt-get update -qq
sudo apt-get upgrade -y -qq

# 2. 安裝 Python 3.11 + pip
echo "[2/7] 安裝 Python..."
sudo apt-get install -y -qq python3 python3-pip python3-venv git

# 3. Clone repo
echo "[3/7] 下載程式碼..."
cd /home/ubuntu
if [ -d "stock-quant-research" ]; then
    cd stock-quant-research
    git pull origin main
else
    git clone https://github.com/cbe566/stock-quant-research.git
    cd stock-quant-research
fi

# 4. 建立虛擬環境 + 安裝依賴
echo "[4/7] 安裝 Python 依賴..."
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip -q
pip install yfinance pandas numpy requests reportlab xlrd PyMuPDF -q

# 5. 建立執行腳本
echo "[5/7] 建立執行腳本..."
cat > /home/ubuntu/run_daily.sh << 'SCRIPT'
#!/bin/bash
# 每日篩選執行腳本
cd /home/ubuntu/stock-quant-research
source venv/bin/activate
export ORACLE_CLOUD=true
export CI=true

echo "$(date '+%Y-%m-%d %H:%M:%S') — 開始每日篩選"

# 更新程式碼
git pull origin main 2>/dev/null || true

# 執行篩選
python3 screening_engine.py 2>&1 | tee logs/screening_$(date +%Y%m%d).log

# 生成 PDF
python3 generate_report_pdf.py 2>&1

# 提交報告
git add daily_reports/ cache/ 2>/dev/null || true
git diff --staged --quiet || {
    git commit -m "📈 每日篩選報告 $(date +%Y-%m-%d) — Oracle Cloud"
    git push origin main 2>/dev/null || true
}

echo "$(date '+%Y-%m-%d %H:%M:%S') — 完成"
SCRIPT
chmod +x /home/ubuntu/run_daily.sh

# 建立 logs 目錄
mkdir -p /home/ubuntu/stock-quant-research/logs

# 6. 設定 cron — 每天早上 06:28 (UTC+8 = 22:28 UTC)
echo "[6/7] 設定 cron 排程..."
(crontab -l 2>/dev/null | grep -v run_daily; echo "28 22 * * * /home/ubuntu/run_daily.sh >> /home/ubuntu/stock-quant-research/logs/cron.log 2>&1") | crontab -

# 7. 設定 git
echo "[7/7] 設定 git..."
cd /home/ubuntu/stock-quant-research
git config user.name "Oracle Cloud Bot"
git config user.email "oracle@stock-quant.auto"

# 完成
echo ""
echo "================================================"
echo "  部署完成！"
echo "================================================"
echo ""
echo "  排程：每天 22:28 UTC (06:28 HKT) 自動執行"
echo "  股票數：4,814 隻（美517 + 港95 + 台1081 + 日3121）"
echo "  日誌：/home/ubuntu/stock-quant-research/logs/"
echo ""
echo "  手動測試："
echo "    bash /home/ubuntu/run_daily.sh"
echo ""
echo "  查看 cron："
echo "    crontab -l"
echo ""
echo "================================================"
