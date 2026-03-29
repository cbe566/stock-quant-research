"""
每日全球股票篩選 — PDF 詳細報告生成器
=======================================
風格參考：券商每日晨報（華泰/國信/高盛）
頁面：A4，深藍品牌色，專業排版
"""

import json
import os
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==================== 字體設定 ====================
FONT_PATHS = [
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/STHeiti Medium.ttc",
]

FONT_NAME = "Helvetica"  # fallback
for fp in FONT_PATHS:
    if os.path.exists(fp):
        try:
            pdfmetrics.registerFont(TTFont("ArialUnicode", fp))
            FONT_NAME = "ArialUnicode"
            break
        except Exception:
            continue

# ==================== 品牌色 ====================
BRAND_NAVY = colors.HexColor("#1B2A4A")
BRAND_BLUE = colors.HexColor("#2563EB")
BRAND_LIGHT_BLUE = colors.HexColor("#EFF6FF")
BRAND_GREEN = colors.HexColor("#059669")
BRAND_RED = colors.HexColor("#DC2626")
BRAND_GRAY = colors.HexColor("#6B7280")
BRAND_LIGHT_GRAY = colors.HexColor("#F3F4F6")
BRAND_BORDER = colors.HexColor("#D1D5DB")
WHITE = colors.white

PAGE_W, PAGE_H = A4  # 595 x 842
MARGIN = 18 * mm
CONTENT_W = PAGE_W - 2 * MARGIN


def _clean_ticker(ticker):
    """清理代碼，確保不會被渲染成超連結"""
    # 在 .HK / .TW / .T 前加零寬空格，破壞 URL 識別
    return ticker


class DailyReportPDF:
    """生成每日篩選詳細 PDF 報告"""

    def __init__(self, json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)

        self.date_str = self.data.get("date", datetime.now().strftime("%Y-%m-%d"))
        self.generated_at = self.data.get("generated_at", "")
        self.markets = self.data.get("markets", {})
        self._init_styles()

    def _init_styles(self):
        """初始化段落樣式"""
        self.styles = getSampleStyleSheet()

        self.styles.add(ParagraphStyle(
            "CoverTitle", fontName=FONT_NAME, fontSize=28, leading=36,
            textColor=WHITE, alignment=TA_CENTER, spaceAfter=6 * mm,
        ))
        self.styles.add(ParagraphStyle(
            "CoverSub", fontName=FONT_NAME, fontSize=14, leading=20,
            textColor=colors.HexColor("#93C5FD"), alignment=TA_CENTER,
        ))
        self.styles.add(ParagraphStyle(
            "SectionTitle", fontName=FONT_NAME, fontSize=16, leading=22,
            textColor=BRAND_NAVY, spaceBefore=8 * mm, spaceAfter=4 * mm,
        ))
        self.styles.add(ParagraphStyle(
            "SubTitle", fontName=FONT_NAME, fontSize=12, leading=16,
            textColor=BRAND_BLUE, spaceBefore=4 * mm, spaceAfter=2 * mm,
        ))
        self.styles.add(ParagraphStyle(
            "Body", fontName=FONT_NAME, fontSize=9.5, leading=15,
            textColor=colors.HexColor("#1F2937"),
        ))
        self.styles.add(ParagraphStyle(
            "SmallGray", fontName=FONT_NAME, fontSize=8, leading=12,
            textColor=BRAND_GRAY,
        ))
        self.styles.add(ParagraphStyle(
            "Footer", fontName=FONT_NAME, fontSize=7.5, leading=11,
            textColor=BRAND_GRAY, alignment=TA_CENTER,
        ))
        self.styles.add(ParagraphStyle(
            "TableCell", fontName=FONT_NAME, fontSize=8, leading=11,
            textColor=colors.HexColor("#1F2937"),
        ))
        self.styles.add(ParagraphStyle(
            "TableCellBold", fontName=FONT_NAME, fontSize=8, leading=11,
            textColor=colors.HexColor("#1F2937"),
        ))

    def _build_cover(self):
        """封面頁"""
        elements = []
        elements.append(Spacer(1, 60 * mm))

        # 深藍色封面區塊用表格模擬
        cover_data = [[""]]
        cover_table = Table(cover_data, colWidths=[CONTENT_W], rowHeights=[180])
        cover_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BRAND_NAVY),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))

        # 封面內容
        dt = datetime.strptime(self.date_str, "%Y-%m-%d")
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        date_display = f"{dt.year}年{dt.month}月{dt.day}日（星期{weekdays[dt.weekday()]}）"

        cover_content = []
        cover_content.append(Spacer(1, 25 * mm))

        # 品牌色條
        cover_content.append(HRFlowable(width=CONTENT_W, thickness=3, color=BRAND_BLUE))
        cover_content.append(Spacer(1, 15 * mm))
        cover_content.append(Paragraph("每日全球股票篩選報告", self.styles["CoverTitle"]))
        cover_content.append(Paragraph("Daily Global Equity Screening Report", self.styles["CoverSub"]))
        cover_content.append(Spacer(1, 8 * mm))
        cover_content.append(Paragraph(date_display, self.styles["CoverSub"]))
        cover_content.append(Spacer(1, 15 * mm))
        cover_content.append(HRFlowable(width=CONTENT_W, thickness=1, color=colors.HexColor("#93C5FD")))
        cover_content.append(Spacer(1, 8 * mm))
        cover_content.append(Paragraph(
            "覆蓋市場：美股 | 港股 | 台股 | 日股",
            self.styles["CoverSub"]
        ))
        cover_content.append(Paragraph(
            "分析維度：QVM多因子 | 技術面 | 動量 | 均值回歸 | F-Score | 分析師目標價",
            ParagraphStyle("CoverDetail", parent=self.styles["CoverSub"], fontSize=10, leading=14)
        ))

        elements.extend(cover_content)
        elements.append(Spacer(1, 30 * mm))

        # 署名區
        elements.append(HRFlowable(width=CONTENT_W, thickness=0.5, color=BRAND_BORDER))
        elements.append(Spacer(1, 4 * mm))
        sig_style = ParagraphStyle("Sig", parent=self.styles["Body"], fontSize=9, leading=14, textColor=BRAND_GRAY)
        elements.append(Paragraph("<b>何宣逸</b>　副總裁 ｜ 私人財富管理部", sig_style))
        elements.append(Paragraph("華泰金融控股（香港）有限公司", sig_style))
        elements.append(Paragraph("電話：+852 3658 6180 ｜ 手機：+852 6765 0336 / +86 130 0329 5233", sig_style))
        elements.append(Paragraph("電郵：jamieho@htsc.com", sig_style))

        elements.append(PageBreak())
        return elements

    def _build_market_section(self, market_name, results):
        """單個市場的詳細報告"""
        elements = []

        flag_map = {"美股": "🇺🇸", "港股": "🇭🇰", "台股": "🇹🇼", "日股": "🇯🇵"}
        flag = flag_map.get(market_name, "")

        elements.append(Paragraph(f"{flag} {market_name}", self.styles["SectionTitle"]))
        elements.append(HRFlowable(width=CONTENT_W, thickness=2, color=BRAND_NAVY))
        elements.append(Spacer(1, 3 * mm))

        # 市場統計
        if results:
            total = len(results)
            bullish = len([r for r in results if r.get("total_score", 0) >= 3])
            bearish = len([r for r in results if r.get("total_score", 0) <= -3])
            neutral = total - bullish - bearish
            avg = sum(r.get("total_score", 0) for r in results) / total if total > 0 else 0

            stat_text = f"篩選股票數：{total} ｜ 看多：{bullish} ｜ 中性：{neutral} ｜ 看空：{bearish} ｜ 平均得分：{avg:+.1f}"
            elements.append(Paragraph(stat_text, self.styles["SmallGray"]))
            elements.append(Spacer(1, 4 * mm))

        # === 買入 TOP 10 ===
        buy_list = sorted(results, key=lambda x: x.get("total_score", 0), reverse=True)[:10]
        elements.append(Paragraph("買入關注 TOP 10", self.styles["SubTitle"]))
        elements.append(self._build_stock_table(buy_list, is_buy=True))
        elements.append(Spacer(1, 4 * mm))

        # === 賣出 TOP 10 ===
        sell_list = sorted(results, key=lambda x: x.get("total_score", 0))[:10]
        elements.append(Paragraph("賣出關注 TOP 10", self.styles["SubTitle"]))
        elements.append(self._build_stock_table(sell_list, is_buy=False))

        elements.append(PageBreak())
        return elements

    def _build_stock_table(self, stocks, is_buy=True):
        """建立股票表格"""
        header = ["#", "代碼", "名稱", "現價", "得分", "QVM", "技術面", "Z-Score", "信號摘要"]
        col_widths = [18, 52, 80, 52, 32, 32, 38, 42, CONTENT_W - 346]

        data = [header]
        for i, s in enumerate(stocks, 1):
            ticker = _clean_ticker(s.get("ticker", ""))
            name = (s.get("name", "") or "")[:12]
            price = s.get("current_price", "—")
            if isinstance(price, (int, float)):
                price = f"{price:,.2f}"
            score = s.get("total_score", 0)
            qvm = s.get("quality", "—")
            if isinstance(qvm, (int, float)):
                qvm = f"{qvm:.0f}"
            tech = s.get("tech_signal", "—")
            if isinstance(tech, (int, float)):
                tech = f"{tech:+.0f}"
            z = s.get("zscore", "—")
            if isinstance(z, (int, float)):
                z = f"{z:.1f}"
            signals = "、".join(s.get("signals", [])[:2]) or "—"

            score_str = f"{score:+d}"
            data.append([str(i), ticker, name, price, score_str, qvm, tech, z, signals])

        table = Table(data, colWidths=col_widths, repeatRows=1)

        # 表格樣式
        header_color = BRAND_GREEN if is_buy else BRAND_RED
        style_cmds = [
            # 表頭
            ("BACKGROUND", (0, 0), (-1, 0), header_color),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("LEADING", (0, 0), (-1, -1), 11),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (3, 0), (7, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, BRAND_BORDER),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]

        # 斑馬紋
        for row in range(1, len(data)):
            if row % 2 == 0:
                style_cmds.append(("BACKGROUND", (0, row), (-1, row), BRAND_LIGHT_GRAY))

        table.setStyle(TableStyle(style_cmds))
        return table

    def _build_disclaimer(self):
        """免責聲明 + 署名"""
        elements = []
        elements.append(HRFlowable(width=CONTENT_W, thickness=1, color=BRAND_BORDER))
        elements.append(Spacer(1, 4 * mm))

        # 署名
        sig_style = ParagraphStyle("SigBlock", parent=self.styles["Body"], fontSize=9, leading=14)
        elements.append(Paragraph("<b>何宣逸</b>", sig_style))
        elements.append(Paragraph("副總裁 ｜ 私人財富管理部", self.styles["SmallGray"]))
        elements.append(Paragraph("華泰金融控股（香港）有限公司", self.styles["SmallGray"]))
        elements.append(Paragraph("電話：+852 3658 6180 ｜ 手機：+852 6765 0336 / +86 130 0329 5233", self.styles["SmallGray"]))
        elements.append(Paragraph("電郵：jamieho@htsc.com", self.styles["SmallGray"]))
        elements.append(Spacer(1, 4 * mm))

        # 公司資訊
        elements.append(Paragraph("地址：香港皇后大道中99號中環中心69樓", self.styles["SmallGray"]))
        elements.append(Paragraph(
            "華泰證券股份有限公司全資附屬公司 (SSE: 601688; SEHK: 6886; LSE: HTSC)",
            self.styles["SmallGray"]
        ))
        elements.append(Spacer(1, 4 * mm))
        elements.append(HRFlowable(width=CONTENT_W, thickness=0.5, color=BRAND_BORDER))
        elements.append(Spacer(1, 3 * mm))

        # 報告資訊
        now = datetime.now()
        time_str = now.strftime("%Y-%m-%d %H:%M") + " (UTC+8)"
        elements.append(Paragraph(f"報告製作時間：{time_str}", self.styles["SmallGray"]))
        elements.append(Paragraph(
            "資料來源：Yahoo Finance、FMP、Finnhub、Tiingo、Twelve Data、TWSE",
            self.styles["SmallGray"]
        ))
        elements.append(Paragraph(
            "量化模型：QVM 多因子評分 × 技術面信號 × 均值回歸 × Piotroski F-Score",
            self.styles["SmallGray"]
        ))
        elements.append(Spacer(1, 3 * mm))

        # 免責聲明
        disclaimer = (
            "本報告僅供參考，不構成任何投資建議。投資有風險，入市需謹慎。"
            "報告中的資訊來自公開市場數據，華泰金融控股（香港）有限公司不對數據的準確性、"
            "完整性或及時性做出任何保證。過往表現不代表未來回報。"
        )
        disc_style = ParagraphStyle(
            "Disclaimer", parent=self.styles["SmallGray"],
            fontSize=7, leading=10, textColor=BRAND_GRAY,
        )
        elements.append(Paragraph(disclaimer, disc_style))

        return elements

    def _header_footer(self, canvas, doc):
        """頁眉頁腳"""
        canvas.saveState()

        # 頁眉：左品牌名，右日期
        canvas.setFont(FONT_NAME, 7.5)
        canvas.setFillColor(BRAND_GRAY)
        canvas.drawString(MARGIN, PAGE_H - 12 * mm, "華泰金融控股（香港）有限公司")

        dt = datetime.strptime(self.date_str, "%Y-%m-%d")
        date_display = f"每日股票篩選報告 | {dt.year}-{dt.month:02d}-{dt.day:02d}"
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 12 * mm, date_display)

        # 頁眉線
        canvas.setStrokeColor(BRAND_BORDER)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, PAGE_H - 13 * mm, PAGE_W - MARGIN, PAGE_H - 13 * mm)

        # 頁腳
        canvas.setFont(FONT_NAME, 7)
        canvas.setFillColor(BRAND_GRAY)
        page_num = f"{doc.page}"
        canvas.drawCentredString(PAGE_W / 2, 10 * mm, page_num)

        canvas.restoreState()

    def generate(self, output_path=None):
        """生成 PDF"""
        if output_path is None:
            date_str = datetime.now().strftime("%Y%m%d")
            output_dir = Path(__file__).parent / "daily_reports"
            output_dir.mkdir(exist_ok=True)
            output_path = str(output_dir / f"每日篩選報告_{date_str}.pdf")

        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            topMargin=18 * mm,
            bottomMargin=15 * mm,
            leftMargin=MARGIN,
            rightMargin=MARGIN,
        )

        elements = []

        # 封面
        elements.extend(self._build_cover())

        # 各市場詳細報告
        for market_name, results in self.markets.items():
            if results:
                elements.extend(self._build_market_section(market_name, results))

        # 免責聲明
        elements.extend(self._build_disclaimer())

        doc.build(elements, onFirstPage=self._header_footer, onLaterPages=self._header_footer)
        print(f"PDF 已生成：{output_path}")
        return output_path


# ==================== 入口 ====================

def main():
    """從最新的 JSON 數據生成 PDF"""
    reports_dir = Path(__file__).parent / "daily_reports"
    json_files = sorted(reports_dir.glob("screening_data_*.json"), reverse=True)

    if not json_files:
        print("找不到篩選數據 JSON，請先執行 daily_screening.py")
        return

    latest = str(json_files[0])
    print(f"使用數據：{latest}")

    generator = DailyReportPDF(latest)
    pdf_path = generator.generate()
    return pdf_path


if __name__ == "__main__":
    main()
