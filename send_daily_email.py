"""
每日篩選報告 — 郵件發送模組
============================
郵件正文：核心摘要（各市場 TOP 3 買入 + TOP 3 賣出 + 市場情緒）
附件：PDF 完整報告

設計原則：
  - 專業券商晨報風格
  - 郵件正文簡潔，一目瞭然
  - 股票代碼防超連結處理
  - 無任何 GitHub / 自動化字眼
  - 底部署名：華泰金融控股
"""

import json
from datetime import datetime
from pathlib import Path


def _safe_ticker(ticker):
    """
    防止 Gmail 把 .HK / .TW / .T 識別成超連結
    方法：在點號前插入零寬空格 &#8203;
    """
    for suffix in [".HK", ".TW", ".T"]:
        if ticker.endswith(suffix):
            return ticker.replace(suffix, "&#8203;" + suffix)
    return ticker


def _score_color(score):
    """得分顏色"""
    if score >= 3:
        return "#059669"  # 綠
    elif score <= -3:
        return "#DC2626"  # 紅
    elif score > 0:
        return "#047857"
    elif score < 0:
        return "#B91C1C"
    return "#6B7280"


def _score_display(score):
    return f'<span style="color:{_score_color(score)};font-weight:bold">{score:+d}</span>'


def generate_email_html(json_path):
    """生成郵件 HTML（核心摘要版）"""

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    date_str = data.get("date", datetime.now().strftime("%Y-%m-%d"))
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    date_display = f"{dt.year}年{dt.month}月{dt.day}日（星期{weekdays[dt.weekday()]}）"

    markets = data.get("markets", {})

    # 開始 HTML
    html = f'''<html>
<head>
<style>
body {{ font-family: -apple-system, 'Helvetica Neue', Arial, 'PingFang SC', sans-serif; color: #1a1a1a; margin: 0; padding: 0; background: #f9fafb; }}
.container {{ max-width: 680px; margin: 0 auto; background: #ffffff; }}
.header {{ background: #1B2A4A; padding: 24px 30px; }}
.header h1 {{ color: #ffffff; font-size: 20px; margin: 0 0 4px 0; font-weight: 600; }}
.header .date {{ color: #93C5FD; font-size: 13px; margin: 0; }}
.body-content {{ padding: 24px 30px; }}
.market-block {{ margin-bottom: 24px; }}
.market-title {{ font-size: 16px; font-weight: 700; color: #1B2A4A; border-bottom: 2px solid #1B2A4A; padding-bottom: 6px; margin-bottom: 12px; }}
.sentiment {{ font-size: 12px; color: #6B7280; margin-bottom: 10px; }}
.direction {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; font-weight: 600; }}
.dir-bull {{ background: #ECFDF5; color: #059669; }}
.dir-bear {{ background: #FEF2F2; color: #DC2626; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 8px; }}
th {{ background: #F3F4F6; color: #374151; font-size: 11px; padding: 6px 8px; text-align: center; border-bottom: 1px solid #D1D5DB; }}
td {{ font-size: 12px; padding: 5px 8px; text-align: center; border-bottom: 1px solid #F3F4F6; }}
td.ticker {{ font-weight: 600; text-align: left; }}
td.name {{ text-align: left; color: #374151; }}
td.signal {{ text-align: left; font-size: 11px; color: #6B7280; }}
.label-buy {{ color: #059669; font-weight: 600; font-size: 12px; margin-bottom: 4px; }}
.label-sell {{ color: #DC2626; font-weight: 600; font-size: 12px; margin-bottom: 4px; }}
.pdf-note {{ background: #EFF6FF; border-radius: 6px; padding: 14px 18px; margin: 20px 0; font-size: 13px; color: #1E40AF; }}
.footer {{ padding: 20px 30px; border-top: 1px solid #E5E7EB; }}
.sig {{ font-size: 12px; color: #374151; line-height: 1.8; }}
.sig strong {{ font-size: 13px; }}
.company {{ font-size: 11px; color: #9CA3AF; line-height: 1.6; margin-top: 12px; }}
.disclaimer {{ font-size: 10px; color: #9CA3AF; line-height: 1.5; margin-top: 12px; padding-top: 12px; border-top: 1px solid #F3F4F6; }}
</style>
</head>
<body>
<div class="container">

<div class="header">
  <h1>每日全球股票篩選報告</h1>
  <p class="date">{date_display} ｜ 美股 · 港股 · 台股 · 日股</p>
</div>

<div class="body-content">
'''

    flag_map = {"美股": "🇺🇸", "港股": "🇭🇰", "台股": "🇹🇼", "日股": "🇯🇵"}

    for market_name, results in markets.items():
        if not results:
            continue

        flag = flag_map.get(market_name, "")
        total = len(results)
        bullish = len([r for r in results if r.get("total_score", 0) >= 3])
        bearish = len([r for r in results if r.get("total_score", 0) <= -3])
        avg = sum(r.get("total_score", 0) for r in results) / total if total > 0 else 0

        if avg > 1:
            sentiment_label = f'<span class="direction dir-bull">偏多 {avg:+.1f}</span>'
        elif avg < -1:
            sentiment_label = f'<span class="direction dir-bear">偏空 {avg:+.1f}</span>'
        else:
            sentiment_label = f'<span class="direction" style="background:#F3F4F6;color:#6B7280">中性 {avg:+.1f}</span>'

        html += f'''
<div class="market-block">
  <div class="market-title">{flag} {market_name}</div>
  <div class="sentiment">篩選 {total} 隻 ｜ 看多 {bullish} · 看空 {bearish} ｜ {sentiment_label}</div>
'''

        # TOP 3 買入
        buy_top3 = sorted(results, key=lambda x: x.get("total_score", 0), reverse=True)[:3]
        html += '  <p class="label-buy">▲ 買入關注</p>\n'
        html += '  <table><tr><th>代碼</th><th>名稱</th><th>現價</th><th>得分</th><th>核心信號</th></tr>\n'
        for s in buy_top3:
            ticker = _safe_ticker(s.get("ticker", ""))
            name = (s.get("name", "") or "")[:12]
            price = s.get("current_price", "—")
            if isinstance(price, (int, float)):
                price = f"{price:,.2f}"
            score = s.get("total_score", 0)
            signals = "、".join(s.get("signals", [])[:2]) or "—"
            html += f'  <tr><td class="ticker">{ticker}</td><td class="name">{name}</td><td>{price}</td><td>{_score_display(score)}</td><td class="signal">{signals}</td></tr>\n'
        html += '  </table>\n'

        # TOP 3 賣出
        sell_top3 = sorted(results, key=lambda x: x.get("total_score", 0))[:3]
        html += '  <p class="label-sell">▼ 賣出關注</p>\n'
        html += '  <table><tr><th>代碼</th><th>名稱</th><th>現價</th><th>得分</th><th>核心信號</th></tr>\n'
        for s in sell_top3:
            ticker = _safe_ticker(s.get("ticker", ""))
            name = (s.get("name", "") or "")[:12]
            price = s.get("current_price", "—")
            if isinstance(price, (int, float)):
                price = f"{price:,.2f}"
            score = s.get("total_score", 0)
            signals = "、".join(s.get("signals", [])[:2]) or "—"
            html += f'  <tr><td class="ticker">{ticker}</td><td class="name">{name}</td><td>{price}</td><td>{_score_display(score)}</td><td class="signal">{signals}</td></tr>\n'
        html += '  </table>\n'

        html += '</div>\n'

    # PDF 附件提示
    html += '''
<div class="pdf-note">
  📎 完整報告（含各市場 TOP 10 買入/賣出詳細分析）請見附件 PDF
</div>
'''

    # Footer 署名
    now = datetime.now()
    time_str = now.strftime("%Y-%m-%d %H:%M") + " (UTC+8)"

    html += f'''
</div>

<div class="footer">
  <div class="sig">
    <strong>何宣逸</strong><br>
    副總裁 ｜ 私人財富管理部<br>
    華泰金融控股（香港）有限公司<br>
    電話：+852 3658 6180 ｜ 手機：+852 6765 0336 / +86 130 0329 5233<br>
    電郵：jamieho@htsc.com
  </div>
  <div class="company">
    地址：香港皇后大道中99號中環中心69樓<br>
    華泰證券股份有限公司全資附屬公司 (SSE: 601688; SEHK: 6886; LSE: HTSC)
  </div>
  <div class="company">
    報告製作時間：{time_str}<br>
    資料來源：Yahoo Finance、FMP、Finnhub、Tiingo、Twelve Data、TWSE<br>
    量化模型：QVM 多因子評分 × 技術面信號 × 均值回歸 × Piotroski F-Score
  </div>
  <div class="disclaimer">
    本報告僅供參考，不構成任何投資建議。投資有風險，入市需謹慎。
  </div>
</div>

</div>
</body>
</html>'''

    return html


def generate_email_subject(json_path):
    """生成郵件標題"""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    date_str = data.get("date", datetime.now().strftime("%Y-%m-%d"))
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"每日全球股票篩選報告 — {dt.year}/{dt.month:02d}/{dt.day:02d}"


def generate_email_plain(json_path):
    """純文字版本"""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    date_str = data.get("date", datetime.now().strftime("%Y-%m-%d"))
    markets = data.get("markets", {})

    lines = [f"每日全球股票篩選報告 — {date_str}", ""]

    for market_name, results in markets.items():
        if not results:
            continue
        lines.append(f"【{market_name}】")
        buy3 = sorted(results, key=lambda x: x.get("total_score", 0), reverse=True)[:3]
        sell3 = sorted(results, key=lambda x: x.get("total_score", 0))[:3]
        lines.append("買入：" + " | ".join(f"{s['ticker']}({s['total_score']:+d})" for s in buy3))
        lines.append("賣出：" + " | ".join(f"{s['ticker']}({s['total_score']:+d})" for s in sell3))
        lines.append("")

    lines.append("完整報告請見附件 PDF")
    lines.append("")
    lines.append("何宣逸 | 副總裁 | 私人財富管理部")
    lines.append("華泰金融控股（香港）有限公司")
    lines.append("本報告僅供參考，不構成任何投資建議。")

    return "\n".join(lines)


if __name__ == "__main__":
    # 測試：生成 HTML 並存檔
    reports_dir = Path(__file__).parent / "daily_reports"
    json_files = sorted(reports_dir.glob("screening_data_*.json"), reverse=True)

    if json_files:
        latest = str(json_files[0])
        html = generate_email_html(latest)
        out = reports_dir / "email_preview.html"
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"郵件預覽已生成：{out}")
        print(f"標題：{generate_email_subject(latest)}")
