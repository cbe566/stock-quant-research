"""
每日全球股票篩選 — PDF 詳細報告生成器 v2
==========================================
專業排版：深藍封面、漸層表頭、圓角統計卡片、分色買賣區塊
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
    PageBreak, HRFlowable, KeepTogether, CondPageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==================== 字體 ====================
FONT_PATHS = [
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/STHeiti Medium.ttc",
]
FONT_NAME = "Helvetica"
for fp in FONT_PATHS:
    if os.path.exists(fp):
        try:
            pdfmetrics.registerFont(TTFont("ArialUnicode", fp))
            FONT_NAME = "ArialUnicode"
            break
        except Exception:
            continue

# ==================== 色彩系統 ====================
C_NAVY = colors.HexColor("#0F172A")
C_DARK_BLUE = colors.HexColor("#1E3A5F")
C_BLUE = colors.HexColor("#2563EB")
C_LIGHT_BLUE = colors.HexColor("#DBEAFE")
C_GREEN = colors.HexColor("#059669")
C_GREEN_BG = colors.HexColor("#ECFDF5")
C_GREEN_HEADER = colors.HexColor("#047857")
C_RED = colors.HexColor("#DC2626")
C_RED_BG = colors.HexColor("#FEF2F2")
C_RED_HEADER = colors.HexColor("#B91C1C")
C_GRAY = colors.HexColor("#6B7280")
C_LIGHT_GRAY = colors.HexColor("#F8FAFC")
C_BORDER = colors.HexColor("#E2E8F0")
C_TEXT = colors.HexColor("#1E293B")
C_TEXT_LIGHT = colors.HexColor("#475569")
WHITE = colors.white

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm
CW = PAGE_W - 2 * MARGIN  # 內容寬度


class DailyReportPDF:

    def __init__(self, json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)
        self.date_str = self.data.get("date", datetime.now().strftime("%Y-%m-%d"))
        self.markets = self.data.get("markets", {})
        self._init_styles()

    def _init_styles(self):
        self.s = {}
        self.s["cover_title"] = ParagraphStyle("ct", fontName=FONT_NAME, fontSize=26, leading=34, textColor=WHITE, alignment=TA_CENTER)
        self.s["cover_sub"] = ParagraphStyle("cs", fontName=FONT_NAME, fontSize=13, leading=18, textColor=colors.HexColor("#94A3B8"), alignment=TA_CENTER)
        self.s["cover_date"] = ParagraphStyle("cd", fontName=FONT_NAME, fontSize=16, leading=22, textColor=colors.HexColor("#93C5FD"), alignment=TA_CENTER)
        self.s["section"] = ParagraphStyle("sec", fontName=FONT_NAME, fontSize=15, leading=20, textColor=C_NAVY, spaceBefore=6*mm, spaceAfter=3*mm)
        self.s["sub_buy"] = ParagraphStyle("sb", fontName=FONT_NAME, fontSize=11, leading=15, textColor=C_GREEN, spaceBefore=4*mm, spaceAfter=2*mm)
        self.s["sub_sell"] = ParagraphStyle("ss", fontName=FONT_NAME, fontSize=11, leading=15, textColor=C_RED, spaceBefore=4*mm, spaceAfter=2*mm)
        self.s["body"] = ParagraphStyle("b", fontName=FONT_NAME, fontSize=9, leading=14, textColor=C_TEXT)
        self.s["small"] = ParagraphStyle("sm", fontName=FONT_NAME, fontSize=7.5, leading=11, textColor=C_GRAY)
        self.s["stat"] = ParagraphStyle("st", fontName=FONT_NAME, fontSize=9, leading=13, textColor=C_TEXT_LIGHT)

    # ==================== 封面 ====================
    def _cover(self):
        el = []
        # 用深藍色大區塊做封面
        dt = datetime.strptime(self.date_str, "%Y-%m-%d")
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        date_disp = f"{dt.year}.{dt.month:02d}.{dt.day:02d}  星期{weekdays[dt.weekday()]}"

        # 計算總股數
        total = sum(len(v) for v in self.markets.values())

        # 封面背景框
        cover_inner = [
            [Paragraph("", ParagraphStyle("sp", fontSize=1))],
        ]
        cover_bg = Table(cover_inner, colWidths=[CW], rowHeights=[220*mm])
        cover_bg.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), C_NAVY),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        # 封面文字（疊在背景上方）
        el.append(Spacer(1, 70*mm))
        el.append(HRFlowable(width=120*mm, thickness=2, color=C_BLUE, spaceAfter=12*mm))
        el.append(Paragraph("每日全球股票篩選報告", self.s["cover_title"]))
        el.append(Spacer(1, 3*mm))
        el.append(Paragraph("Daily Global Equity Screening Report", self.s["cover_sub"]))
        el.append(Spacer(1, 12*mm))
        el.append(HRFlowable(width=60*mm, thickness=0.5, color=colors.HexColor("#475569"), spaceAfter=8*mm))
        el.append(Paragraph(date_disp, self.s["cover_date"]))
        el.append(Spacer(1, 10*mm))
        el.append(Paragraph(f"篩選範圍：{total} 隻股票  |  美股 · 港股 · 台股 · 日股", self.s["cover_sub"]))
        el.append(Spacer(1, 4*mm))
        el.append(Paragraph("QVM多因子 · 技術面 · 動量 · 均值回歸 · F-Score · 分析師目標價",
                            ParagraphStyle("cd2", parent=self.s["cover_sub"], fontSize=10, leading=14)))
        el.append(Spacer(1, 30*mm))

        # 署名
        el.append(HRFlowable(width=CW, thickness=0.5, color=C_BORDER))
        el.append(Spacer(1, 4*mm))
        sig = ParagraphStyle("sig", fontName=FONT_NAME, fontSize=9, leading=14, textColor=C_GRAY)
        el.append(Paragraph("<b>何宣逸 Jamie Ho</b>", sig))
        el.append(Paragraph("手機：+852 6765 0336 / +86 130 0329 5233 ｜ 電郵：cbe566@gmail.com", sig))

        el.append(PageBreak())
        return el

    # ==================== 市場統計卡片 ====================
    def _stat_card(self, results):
        """統計數據卡片"""
        total = len(results)
        bullish = len([r for r in results if r.get("total_score", 0) >= 3])
        bearish = len([r for r in results if r.get("total_score", 0) <= -3])
        avg = sum(r.get("total_score", 0) for r in results) / total if total else 0

        data = [[
            f"篩選：{total} 隻",
            f"看多：{bullish}",
            f"看空：{bearish}",
            f"平均：{avg:+.1f}",
        ]]
        t = Table(data, colWidths=[CW/4]*4, rowHeights=[22])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), C_LIGHT_BLUE),
            ("TEXTCOLOR", (0, 0), (-1, -1), C_DARK_BLUE),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOX", (0, 0), (-1, -1), 0.5, C_BLUE),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    # ==================== 股票表格 ====================
    def _stock_table(self, stocks, is_buy):
        header = ["#", "代碼", "名稱", "現價", "得分", "QVM", "技術面", "Z-Score", "信號摘要"]
        col_w = [16, 50, 78, 50, 30, 30, 36, 40, CW - 330]

        data = [header]
        for i, s in enumerate(stocks[:10], 1):
            ticker = s.get("ticker", "")
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
            data.append([str(i), ticker, name, price, f"{score:+d}", str(qvm), str(tech), str(z), signals])

        t = Table(data, colWidths=col_w, repeatRows=1)

        hdr_color = C_GREEN_HEADER if is_buy else C_RED_HEADER
        row_bg = C_GREEN_BG if is_buy else C_RED_BG

        cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), hdr_color),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("LEADING", (0, 0), (-1, -1), 11),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (3, 0), (7, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.3, C_BORDER),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            # 表頭底線加粗
            ("LINEBELOW", (0, 0), (-1, 0), 1.5, hdr_color),
        ]
        # 交替底色
        for row in range(1, len(data)):
            if row % 2 == 0:
                cmds.append(("BACKGROUND", (0, row), (-1, row), row_bg))

        t.setStyle(TableStyle(cmds))
        return t

    # ==================== 市場章節 ====================
    def _market_section(self, name, results):
        el = []

        # 標題 + 底線
        el.append(Paragraph(f"{name}", self.s["section"]))
        el.append(HRFlowable(width=CW, thickness=2, color=C_NAVY, spaceAfter=3*mm))

        # 統計卡片
        el.append(self._stat_card(results))
        el.append(Spacer(1, 3*mm))

        # 買入 TOP 10（只取得分 > 0 的）
        buy = [r for r in sorted(results, key=lambda x: x.get("total_score", 0), reverse=True) if r.get("total_score", 0) > 0][:10]
        if buy:
            el.append(Paragraph("▲ 買入關注 TOP 10", self.s["sub_buy"]))
            el.append(self._stock_table(buy, is_buy=True))
            el.append(Spacer(1, 3*mm))

        # 賣出 TOP 10（只取得分 < 0 的）
        sell = [r for r in sorted(results, key=lambda x: x.get("total_score", 0)) if r.get("total_score", 0) < 0][:10]
        if sell:
            el.append(Paragraph("▼ 賣出關注 TOP 10", self.s["sub_sell"]))
            el.append(self._stock_table(sell, is_buy=False))

        el.append(PageBreak())
        return el

    # ==================== 免責聲明 ====================
    def _disclaimer(self):
        el = []
        el.append(Spacer(1, 6*mm))
        el.append(HRFlowable(width=CW, thickness=1, color=C_BORDER, spaceAfter=4*mm))

        sig = ParagraphStyle("sig2", fontName=FONT_NAME, fontSize=9, leading=14, textColor=C_TEXT)
        el.append(Paragraph("<b>何宣逸 Jamie Ho</b>", sig))
        sm = self.s["small"]
        el.append(Paragraph("手機：+852 6765 0336 / +86 130 0329 5233 ｜ 電郵：cbe566@gmail.com", sm))
        el.append(Spacer(1, 4*mm))

        now = datetime.now()
        el.append(Paragraph(f"報告製作時間：{now.strftime('%Y-%m-%d %H:%M')} (UTC+8)", sm))
        el.append(Paragraph("資料來源：Yahoo Finance、FMP、Finnhub、Tiingo、Twelve Data、TWSE", sm))
        el.append(Paragraph("量化模型：QVM 多因子評分 x 技術面信號 x 均值回歸 x Piotroski F-Score", sm))
        el.append(Spacer(1, 3*mm))

        disc = ParagraphStyle("disc", fontName=FONT_NAME, fontSize=7, leading=10, textColor=C_GRAY)
        el.append(Paragraph(
            "本報告僅供參考，不構成任何投資建議。投資有風險，入市需謹慎。"
            "報告中的資訊來自公開市場數據，作者不對數據的準確性、完整性或及時性做出任何保證。過往表現不代表未來回報。",
            disc
        ))
        return el

    # ==================== 頁眉頁腳 ====================
    def _header_footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 7)
        canvas.setFillColor(C_GRAY)
        canvas.drawString(MARGIN, PAGE_H - 12*mm, "何宣逸 | 量化研究")

        dt = datetime.strptime(self.date_str, "%Y-%m-%d")
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 12*mm, f"每日股票篩選報告 | {dt.year}-{dt.month:02d}-{dt.day:02d}")

        canvas.setStrokeColor(C_BORDER)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, PAGE_H - 13*mm, PAGE_W - MARGIN, PAGE_H - 13*mm)

        # 頁碼
        canvas.drawCentredString(PAGE_W/2, 10*mm, str(doc.page))
        canvas.restoreState()

    # ==================== 生成 ====================
    def generate(self, output_path=None):
        if output_path is None:
            d = datetime.now().strftime("%Y%m%d")
            out_dir = Path(__file__).parent / "daily_reports"
            out_dir.mkdir(exist_ok=True)
            output_path = str(out_dir / f"每日篩選報告_{d}.pdf")

        doc = SimpleDocTemplate(output_path, pagesize=A4,
                                topMargin=18*mm, bottomMargin=15*mm,
                                leftMargin=MARGIN, rightMargin=MARGIN)

        el = []
        el.extend(self._cover())

        for name, results in self.markets.items():
            if results:
                el.extend(self._market_section(name, results))

        el.extend(self._disclaimer())
        doc.build(el, onFirstPage=self._header_footer, onLaterPages=self._header_footer)
        print(f"PDF 已生成：{output_path}")
        return output_path


def main():
    reports_dir = Path(__file__).parent / "daily_reports"
    json_files = sorted(reports_dir.glob("screening_data_*.json"), reverse=True)
    if not json_files:
        print("找不到篩選數據 JSON")
        return
    latest = str(json_files[0])
    print(f"使用數據：{latest}")
    gen = DailyReportPDF(latest)
    return gen.generate()


if __name__ == "__main__":
    main()
