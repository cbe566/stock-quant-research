# 專業金融投資 PDF 報告製作 — Python 完整知識體系

> 最後更新：2026-03-28
> 適用環境：Python 3.10+、ReportLab 4.x、WeasyPrint 68.x、FPDF2 2.8.x

---

## 目錄

1. [Python PDF 生成工具比較](#一python-pdf-生成工具比較)
2. [專業金融 PDF 排版設計](#二專業金融-pdf-排版設計)
3. [ReportLab 完整教程](#三reportlab-完整教程)
4. [WeasyPrint（HTML→PDF）教程](#四weasyprint-htmlpdf-教程)
5. [完整代碼範例](#五完整代碼範例)
6. [自動化報告流程](#六自動化報告流程)

---

## 一、Python PDF 生成工具比較

### 1.1 主要工具總覽

| 工具 | 類型 | 學習曲線 | 中文支持 | 效能(1頁) | 適合場景 |
|------|------|---------|---------|----------|---------|
| **ReportLab** | 原生 PDF 生成 | 陡峭 | 優秀（CID/TTF） | ~0.08s | 複雜金融報告、精密排版 |
| **FPDF2** | 輕量 PDF 生成 | 簡單 | 良好（Unicode） | ~0.05s | 簡單報告、快速原型 |
| **WeasyPrint** | HTML/CSS → PDF | 中等 | 良好（系統字體） | ~0.35s | 設計驅動報告、模板化 |
| **xhtml2pdf** | HTML → PDF | 簡單 | 一般 | ~0.20s | 簡單 HTML 轉換 |
| **Playwright** | 瀏覽器渲染 → PDF | 中等 | 優秀 | ~0.75s | 高保真網頁轉換 |
| **borb** | PDF 操作 | 中等 | 一般 | ~0.12s | PDF 編輯、互動表單 |
| **matplotlib** | 圖表 → PDF | 簡單 | 一般 | 依圖表 | 單純圖表輸出 |
| **Plotly** | 互動圖表 → 靜態 | 簡單 | 一般 | 依圖表 | 精美互動圖表 |

### 1.2 各工具詳細分析

#### ReportLab — 金融報告首選

**優點：**
- Platypus 高級排版引擎，支持自動分頁、多欄佈局
- 原生圖表引擎（柱狀圖、折線圖、圓餅圖）
- 完整的中文字體支持（Adobe CID 字體 + TrueType）
- 業界標準，大量企業級應用（發票、財報、研究報告）
- 精確的座標控制（Canvas API）
- 支持 PDF 加密、數位簽章
- 表格功能強大（合併、跨頁、條件格式）

**缺點：**
- 學習曲線較陡，API 複雜
- 排版需要程式碼控制，無法所見即所得
- 樣式調整需大量程式碼

**安裝：**
```bash
pip install reportlab
```

#### FPDF2 — 輕量快速

**優點：**
- 無外部依賴，安裝極簡
- API 直觀，上手快
- 生成速度最快（~0.05 秒/頁）
- Unicode 支持良好
- 適合批量快速生成

**缺點：**
- 排版能力有限，無自動分頁引擎
- 不支持 HTML/CSS
- 圖表需外部生成後嵌入
- 複雜表格支持不如 ReportLab

**安裝：**
```bash
pip install fpdf2
```

#### WeasyPrint — 設計師友好

**優點：**
- 用 HTML + CSS 設計，對前端開發者友好
- 支持 CSS Grid、Flexbox、@page 規則
- 配合 Jinja2 實現模板化報告
- 樣式修改直觀（改 CSS 即可）
- 支持 CSS 分頁控制

**缺點：**
- 依賴系統級庫（cairo、pango、gdk-pixbuf）
- 安裝較複雜（macOS 需 brew 安裝依賴）
- 某些 CSS 特性不完全支持
- 不支持 PDF 表單、無障礙 PDF
- 效能較慢

**安裝：**
```bash
# macOS
brew install pango gdk-pixbuf libffi
pip install weasyprint

# Ubuntu/Debian
sudo apt install libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0
pip install weasyprint
```

#### Jinja2 + WeasyPrint — 模板化方案

**優點：**
- 模板繼承與複用
- 動態數據填入（循環、條件判斷）
- HTML 預覽後再轉 PDF
- 設計與數據分離

**缺點：**
- 兩步流程（渲染 HTML → 轉 PDF）
- 偵錯需檢查 HTML 中間產出

### 1.3 最佳方案推薦

#### 金融報告最佳選擇：ReportLab

**理由：**
1. **精密排版控制** — 金融報告需要精確的表格對齊、數字格式、條件配色
2. **原生圖表** — 無需依賴外部圖表庫（但也可嵌入 matplotlib）
3. **中文支持最完善** — Adobe CID 字體不需嵌入，檔案更小
4. **業界標準** — 銀行、券商、基金公司廣泛使用
5. **PDF 功能完整** — 加密、書籤、超連結、元數據

#### 設計導向報告：WeasyPrint + Jinja2

**適合場景：**
- 團隊中有前端設計師
- 需要快速迭代視覺設計
- 報告樣式需經常變更
- 已有 HTML 模板資源

#### 快速原型/簡單報告：FPDF2

**適合場景：**
- 簡單數據摘要
- 批量生成大量相似報告
- 伺服器資源有限

#### 混合方案（推薦）

```
最佳實踐：ReportLab（排版引擎） + matplotlib/Plotly（圖表） + pandas（數據）
```

此組合是業界標準的金融報告自動化技術棧。

---

## 二、專業金融 PDF 排版設計

### 2.1 頂級投行報告設計規範

基於高盛（Goldman Sachs）、摩根士丹利（Morgan Stanley）、摩根大通（J.P. Morgan）等投行研究報告的設計分析：

#### 頁面基礎設定

| 項目 | 標準設定 |
|------|---------|
| 頁面尺寸 | **A4**（210mm x 297mm），美國投行偶用 Letter（8.5" x 11"） |
| 上邊距 | 25-30mm（留給頁首 Logo 與分隔線） |
| 下邊距 | 20-25mm（留給頁尾免責聲明與頁碼） |
| 左邊距 | 20-25mm |
| 右邊距 | 15-20mm |
| 正文字體 | 英文：Arial / Helvetica / Calibri，10-11pt |
| 中文字體 | 微軟正黑 / 思源黑體 / 蘋方，10-11pt |
| 行距 | 1.2-1.4 倍行距 |

#### 封面設計

```
┌─────────────────────────────────────┐
│ [公司Logo]           [日期/編號]     │
│                                      │
│                                      │
│     ┌─────────────────────────┐      │
│     │   股票研究報告            │      │
│     │   XXXX 公司（代碼）      │      │
│     │                          │      │
│     │   投資評級：買入 ▲        │      │
│     │   目標價：$XXX           │      │
│     │   現價：$XXX             │      │
│     │   潛在漲幅：XX%          │      │
│     └─────────────────────────┘      │
│                                      │
│     [股價走勢迷你圖]                  │
│                                      │
│     分析師：XXX                       │
│     團隊：XXX                        │
│     郵箱：xxx@xxx.com                │
│                                      │
│     [免責聲明摘要]                    │
└─────────────────────────────────────┘
```

#### 頁首設計

```
┌─────────────────────────────────────┐
│ [Logo] │ 股票研究 │ 2026年3月28日   │
│─────────────────────────────────────│
```

- 左側：公司 Logo
- 中間：報告類別（股票研究 / 宏觀日報 / 策略周報）
- 右側：日期
- 底部：分隔線（通常 0.5pt，品牌色）

#### 頁尾設計

```
│─────────────────────────────────────│
│ 免責聲明文字（6-7pt灰色）     第X頁  │
└─────────────────────────────────────┘
```

#### 標題層次

| 層級 | 字體 | 大小 | 顏色 | 樣式 |
|------|------|------|------|------|
| H1（章節標題） | 粗體無襯線 | 16-18pt | 品牌深藍 | 下方加分隔線 |
| H2（次標題） | 粗體無襯線 | 13-14pt | 品牌深藍 | 無分隔線 |
| H3（小標題） | 粗體無襯線 | 11-12pt | 深灰/品牌色 | 可加項目符號 |
| 正文 | 常規無襯線 | 10-11pt | 黑色/深灰 | — |
| 表格標題 | 粗體 | 9-10pt | 白色(深色底) | 居中 |
| 表格內容 | 常規 | 8-9pt | 深灰 | 數字右對齊 |
| 腳注 | 常規 | 7-8pt | 灰色 | — |

#### 分欄佈局

- **研究報告**：通常單欄，寬度適中便於閱讀
- **宏觀日報**：可用雙欄，左側文字右側圖表
- **數據密集頁**：可用三欄（如多個指標並排）
- **摘要框**：獨立的重點摘要框，淺色背景+邊框

### 2.2 專業配色方案

#### 投行風格配色

```python
# 主色調（深藍色系 — 信任、專業）
PRIMARY_DARK = '#1B2A4A'      # 深海軍藍 — 標題、頁首
PRIMARY_MID = '#2D5F8A'       # 中藍 — 次標題、強調
PRIMARY_LIGHT = '#4A90D9'     # 亮藍 — 連結、互動
PRIMARY_PALE = '#E8F0FE'      # 淡藍 — 背景、高亮

# 輔助色
ACCENT_GOLD = '#C4A35A'       # 金色 — 重要標記
ACCENT_GREEN = '#2E7D32'      # 深綠 — 正面/上漲
ACCENT_RED = '#C62828'        # 深紅 — 負面/下跌

# 灰階
GRAY_900 = '#212121'          # 正文
GRAY_700 = '#616161'          # 次要文字
GRAY_500 = '#9E9E9E'          # 輔助文字
GRAY_300 = '#E0E0E0'          # 分隔線
GRAY_100 = '#F5F5F5'          # 表格條紋背景
WHITE = '#FFFFFF'             # 白色背景
```

#### 表格配色規範

```python
# 表格樣式
TABLE_HEADER_BG = '#1B2A4A'   # 深藍表頭背景
TABLE_HEADER_TEXT = '#FFFFFF'  # 白色表頭文字
TABLE_ROW_ODD = '#FFFFFF'     # 奇數行白色
TABLE_ROW_EVEN = '#F5F7FA'    # 偶數行淡藍灰
TABLE_BORDER = '#D0D5DD'      # 表格邊框淡灰
TABLE_HIGHLIGHT = '#FFF8E1'   # 高亮行淡黃

# 漲跌配色
COLOR_UP = '#2E7D32'          # 上漲深綠
COLOR_DOWN = '#C62828'        # 下跌深紅
COLOR_FLAT = '#616161'        # 持平灰色
```

#### 圖表配色

```python
# 圖表色盤（最多8色，依序使用）
CHART_PALETTE = [
    '#2D5F8A',  # 主藍
    '#C4A35A',  # 金色
    '#2E7D32',  # 深綠
    '#C62828',  # 深紅
    '#6A1B9A',  # 紫色
    '#E65100',  # 橘色
    '#00838F',  # 青色
    '#4E342E',  # 棕色
]
```

#### 印刷 vs 螢幕注意事項

| 考量 | 螢幕顯示 | 印刷輸出 |
|------|---------|---------|
| 色彩模式 | RGB | 建議 CMYK（ReportLab 支持） |
| 對比度 | 至少 4.5:1 | 至少 7:1（紙張吸墨） |
| 細線條 | 0.5pt 可見 | 至少 0.75pt |
| 淺色背景 | 可用 10% 透明度 | 至少 15% 飽和度 |
| 漸層 | 自由使用 | 避免過度漸層（印刷易色帶） |

### 2.3 圖表在 PDF 中的最佳實踐

#### matplotlib 高品質輸出設定

```python
import matplotlib.pyplot as plt
import matplotlib

# 全域設定
matplotlib.rcParams.update({
    'figure.dpi': 150,           # 螢幕預覽 DPI
    'savefig.dpi': 300,          # 儲存 DPI（印刷品質）
    'figure.figsize': (8, 5),    # 預設尺寸（英寸）
    'font.family': 'sans-serif',
    'font.sans-serif': ['Microsoft JhengHei', 'Noto Sans TC', 'Arial'],
    'axes.unicode_minus': False,  # 修復負號顯示
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'savefig.transparent': False, # PDF 中通常不用透明背景
    'savefig.bbox': 'tight',
    'savefig.pad_inches': 0.1,
})
```

#### 向量圖 vs 位圖選擇

| 格式 | 優點 | 缺點 | 適用場景 |
|------|------|------|---------|
| **SVG → PDF** | 無限縮放、檔案小 | 複雜圖表渲染慢 | 折線圖、柱狀圖 |
| **PNG（300dpi）** | 通用性好 | 檔案較大、放大失真 | 含大量數據點的圖 |
| **PDF 原生** | 最佳品質 | 僅 ReportLab 原生支持 | ReportLab 圖表 |

#### 圖表嵌入最佳方式（BytesIO 記憶體方法）

```python
import io
from reportlab.lib.utils import ImageReader

def create_chart_image(fig):
    """將 matplotlib 圖表轉為 ReportLab 可用的記憶體圖像"""
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close(fig)
    buffer.seek(0)
    return ImageReader(buffer)
```

#### 字體嵌入確認

```python
# matplotlib 嵌入字體到 PDF
matplotlib.rcParams['pdf.fonttype'] = 42  # TrueType 字體嵌入
matplotlib.rcParams['ps.fonttype'] = 42

# 確認字體可用
from matplotlib.font_manager import FontManager
fm = FontManager()
available_fonts = [f.name for f in fm.ttflist]
# 檢查中文字體是否在列表中
```

---

## 三、ReportLab 完整教程

### 3.1 基礎篇

#### 安裝與驗證

```bash
pip install reportlab
python -c "import reportlab; print(reportlab.Version)"
# 預期輸出：4.4.10（或更新版本）
```

#### 中文字體配置

```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# ========== 方法一：Adobe CID 字體（不嵌入，檔案最小） ==========
# 簡體中文
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
# 繁體中文
pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))    # 宋體
pdfmetrics.registerFont(UnicodeCIDFont('MHei-Medium'))    # 黑體

# ========== 方法二：TrueType 字體（嵌入字體子集） ==========
# macOS 系統字體路徑
pdfmetrics.registerFont(TTFont('MicrosoftJhengHei',
    '/System/Library/Fonts/STHeiti Medium.ttc'))

# 或使用思源黑體（需先下載）
pdfmetrics.registerFont(TTFont('NotoSansTC',
    '/path/to/NotoSansTC-Regular.otf'))
pdfmetrics.registerFont(TTFont('NotoSansTC-Bold',
    '/path/to/NotoSansTC-Bold.otf'))

# ========== 方法三：自動偵測系統中文字體 ==========
import os
import platform

def find_chinese_font():
    """自動偵測系統中的中文字體"""
    system = platform.system()
    font_paths = []

    if system == 'Darwin':  # macOS
        font_paths = [
            '/System/Library/Fonts/STHeiti Medium.ttc',
            '/System/Library/Fonts/PingFang.ttc',
            '/Library/Fonts/Arial Unicode.ttf',
        ]
    elif system == 'Windows':
        font_paths = [
            'C:/Windows/Fonts/msyh.ttc',      # 微軟雅黑
            'C:/Windows/Fonts/msjh.ttc',       # 微軟正黑
            'C:/Windows/Fonts/simsun.ttc',     # 宋體
        ]
    elif system == 'Linux':
        font_paths = [
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
        ]

    for path in font_paths:
        if os.path.exists(path):
            return path

    return None

# 使用
font_path = find_chinese_font()
if font_path:
    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
```

#### Canvas 基礎操作

```python
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm, inch
from reportlab.lib.colors import HexColor

# 建立畫布
c = canvas.Canvas("基礎範例.pdf", pagesize=A4)
width, height = A4  # 595.27, 841.89 點（1點 = 1/72英寸）

# 繪製文字
c.setFont('Helvetica-Bold', 24)
c.drawString(2*cm, height - 3*cm, "標題文字")

c.setFont('Helvetica', 12)
c.drawString(2*cm, height - 5*cm, "正文內容")

# 繪製線條
c.setStrokeColor(HexColor('#1B2A4A'))
c.setLineWidth(1)
c.line(2*cm, height - 3.5*cm, width - 2*cm, height - 3.5*cm)

# 繪製矩形（填充）
c.setFillColor(HexColor('#E8F0FE'))
c.rect(2*cm, height - 8*cm, width - 4*cm, 2*cm, fill=1, stroke=0)

# 繪製圖片
# c.drawImage("logo.png", 2*cm, height - 2*cm, width=3*cm, height=1*cm)

# 儲存
c.showPage()
c.save()
```

### 3.2 Platypus 高級排版引擎

Platypus（Page Layout and Typography Using Scripts）是 ReportLab 的高級排版框架，自動處理分頁、文字流動等複雜排版。

#### 核心概念架構

```
BaseDocTemplate
  └── PageTemplate（定義頁面佈局）
        └── Frame（定義內容區域）
              └── Flowable（內容元素）
                    ├── Paragraph（段落）
                    ├── Table（表格）
                    ├── Image（圖片）
                    ├── Spacer（間距）
                    ├── PageBreak（分頁）
                    ├── KeepTogether（不分頁群組）
                    └── 自定義 Flowable
```

#### Story（故事流）

```python
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

# Story 是一個 Flowable 物件的列表，依序排入 Frame
story = []

# 取得預設樣式
styles = getSampleStyleSheet()

# 加入內容
story.append(Paragraph("第一章：市場概覽", styles['Heading1']))
story.append(Spacer(1, 0.5*cm))
story.append(Paragraph("本報告分析全球市場...", styles['Normal']))
story.append(PageBreak())  # 強制分頁
story.append(Paragraph("第二章：個股分析", styles['Heading1']))
```

#### 自定義段落樣式

```python
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

# 標題樣式
style_h1 = ParagraphStyle(
    name='CustomH1',
    fontName='Helvetica-Bold',       # 或 'NotoSansTC-Bold' 用中文
    fontSize=18,
    leading=22,                       # 行距（字體大小 + 4pt）
    textColor=HexColor('#1B2A4A'),
    spaceAfter=12,                    # 段後間距
    spaceBefore=20,                   # 段前間距
    alignment=TA_LEFT,
    borderWidth=0,
    borderColor=HexColor('#1B2A4A'),
    borderPadding=(0, 0, 3, 0),      # 上右下左
)

# 正文樣式
style_body = ParagraphStyle(
    name='CustomBody',
    fontName='Helvetica',
    fontSize=10,
    leading=14,
    textColor=HexColor('#212121'),
    spaceAfter=6,
    alignment=TA_JUSTIFY,             # 兩端對齊
    firstLineIndent=0,                # 首行縮排
)

# 重點框樣式
style_highlight = ParagraphStyle(
    name='Highlight',
    fontName='Helvetica',
    fontSize=10,
    leading=14,
    textColor=HexColor('#1B2A4A'),
    backColor=HexColor('#E8F0FE'),
    borderWidth=1,
    borderColor=HexColor('#4A90D9'),
    borderPadding=10,
    spaceAfter=12,
)

# 表格標題樣式
style_table_title = ParagraphStyle(
    name='TableTitle',
    fontName='Helvetica-Bold',
    fontSize=9,
    leading=11,
    textColor=HexColor('#616161'),
    spaceAfter=4,
)
```

#### 段落中的 HTML 標記

```python
# Paragraph 支持有限的 HTML 標記
text = """
<b>粗體文字</b> 和 <i>斜體文字</i><br/>
<font color="#C62828">紅色警告</font> 和
<font color="#2E7D32">綠色正面</font><br/>
<a href="https://example.com">超連結</a><br/>
上標：E=mc<super>2</super>，下標：H<sub>2</sub>O<br/>
<u>底線文字</u> 和 <strike>刪除線</strike>
"""
story.append(Paragraph(text, styles['Normal']))
```

#### Table（表格）完整教程

```python
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor

# ========== 基本表格 ==========
data = [
    ['指標', '2024', '2025E', '2026E', 'YoY%'],
    ['營收(億)', '1,250', '1,450', '1,680', '+15.9%'],
    ['毛利率', '42.3%', '43.1%', '44.0%', '+0.9pp'],
    ['淨利(億)', '180', '215', '260', '+20.9%'],
    ['EPS', '5.20', '6.22', '7.51', '+20.7%'],
    ['P/E', '18.5x', '15.5x', '12.8x', '—'],
]

table = Table(data, colWidths=[80, 70, 70, 70, 60])

# ========== TableStyle 完整命令參考 ==========
style = TableStyle([
    # 表頭背景與文字
    ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1B2A4A')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 9),
    ('ALIGNMENT', (0, 0), (-1, 0), 'CENTER'),

    # 資料行樣式
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 1), (-1, -1), 8),
    ('TEXTCOLOR', (0, 1), (-1, -1), HexColor('#212121')),

    # 第一欄（指標名）左對齊，其餘右對齊
    ('ALIGNMENT', (0, 1), (0, -1), 'LEFT'),
    ('ALIGNMENT', (1, 1), (-1, -1), 'RIGHT'),

    # 條紋背景
    ('ROWBACKGROUNDS', (0, 1), (-1, -1),
     [HexColor('#FFFFFF'), HexColor('#F5F7FA')]),

    # 邊框
    ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#D0D5DD')),
    ('LINEBELOW', (0, 0), (-1, 0), 1.5, HexColor('#1B2A4A')),
    ('LINEBELOW', (0, -1), (-1, -1), 1, HexColor('#1B2A4A')),

    # 內邊距
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),

    # 垂直對齊
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
])

table.setStyle(style)
story.append(table)
```

#### TableStyle 完整命令清單

```python
# ===== 背景與顏色 =====
('BACKGROUND', (col, row), (col, row), color)        # 單元格背景
('TEXTCOLOR', (col, row), (col, row), color)          # 文字顏色
('ROWBACKGROUNDS', (0, start), (-1, end), [c1, c2])   # 交替行背景

# ===== 字體 =====
('FONTNAME', (col, row), (col, row), 'FontName')      # 字體名
('FONTSIZE', (col, row), (col, row), size)             # 字體大小
('LEADING', (col, row), (col, row), leading)           # 行距

# ===== 對齊 =====
('ALIGNMENT', (col, row), (col, row), 'LEFT/CENTER/RIGHT')  # 水平對齊
('VALIGN', (col, row), (col, row), 'TOP/MIDDLE/BOTTOM')     # 垂直對齊

# ===== 邊框線條 =====
('GRID', (col, row), (col, row), width, color)        # 完整網格
('BOX', (col, row), (col, row), width, color)          # 外框
('OUTLINE', (col, row), (col, row), width, color)      # 同 BOX
('INNERGRID', (col, row), (col, row), width, color)    # 內部網格
('LINEBELOW', (col, row), (col, row), width, color)    # 下方線
('LINEABOVE', (col, row), (col, row), width, color)    # 上方線
('LINEBEFORE', (col, row), (col, row), width, color)   # 左方線
('LINEAFTER', (col, row), (col, row), width, color)    # 右方線

# ===== 間距 =====
('TOPPADDING', (col, row), (col, row), points)         # 上內邊距
('BOTTOMPADDING', (col, row), (col, row), points)      # 下內邊距
('LEFTPADDING', (col, row), (col, row), points)        # 左內邊距
('RIGHTPADDING', (col, row), (col, row), points)       # 右內邊距

# ===== 合併 =====
('SPAN', (start_col, start_row), (end_col, end_row))   # 合併儲存格

# ===== 特殊 =====
('NOSPLIT', (0, row), (-1, row))                        # 禁止分頁
```

#### 條件格式（漲跌配色）

```python
def apply_conditional_format(table_data, table_obj, col_index, start_row=1):
    """根據數值正負自動配色"""
    style_commands = []
    for row_idx in range(start_row, len(table_data)):
        value = table_data[row_idx][col_index]
        # 嘗試解析數值
        try:
            clean_val = str(value).replace('%', '').replace('+', '').replace(',', '')
            num = float(clean_val)
            if num > 0:
                color = HexColor('#2E7D32')  # 綠色（上漲）
            elif num < 0:
                color = HexColor('#C62828')  # 紅色（下跌）
            else:
                color = HexColor('#616161')  # 灰色（持平）
            style_commands.append(
                ('TEXTCOLOR', (col_index, row_idx), (col_index, row_idx), color)
            )
        except (ValueError, TypeError):
            pass

    if style_commands:
        table_obj.setStyle(TableStyle(style_commands))
```

#### Image（圖片嵌入）

```python
from reportlab.platypus import Image
from reportlab.lib.units import cm, inch

# 從檔案嵌入
img = Image('chart.png', width=15*cm, height=9*cm)
story.append(img)

# 從記憶體嵌入（BytesIO）
import io
from reportlab.lib.utils import ImageReader

buffer = io.BytesIO()
fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
plt.close(fig)
buffer.seek(0)

img = Image(ImageReader(buffer), width=15*cm, height=9*cm)
story.append(img)

# 保持比例縮放
img = Image('chart.png')
img_width = 15*cm
ratio = img_width / img.imageWidth
img_height = img.imageHeight * ratio
img = Image('chart.png', width=img_width, height=img_height)
```

#### 間距與分頁控制

```python
from reportlab.platypus import Spacer, PageBreak, CondPageBreak, KeepTogether

# 固定間距
story.append(Spacer(1, 1*cm))

# 強制分頁
story.append(PageBreak())

# 條件分頁：剩餘空間不足 5cm 時分頁
story.append(CondPageBreak(5*cm))

# 保持在同一頁（標題+內容不分離）
story.append(KeepTogether([
    Paragraph("重要圖表", styles['Heading2']),
    Spacer(1, 0.3*cm),
    Image('chart.png', width=15*cm, height=9*cm),
]))
```

#### 自定義 Flowable

```python
from reportlab.platypus import Flowable
from reportlab.lib.colors import HexColor

class HorizontalLine(Flowable):
    """自定義水平分隔線"""
    def __init__(self, width, color='#D0D5DD', thickness=0.5):
        Flowable.__init__(self)
        self.width = width
        self.color = HexColor(color)
        self.thickness = thickness

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

    def wrap(self, availWidth, availHeight):
        self.width = min(self.width, availWidth)
        return (self.width, self.thickness + 2)

class HighlightBox(Flowable):
    """重點提示框"""
    def __init__(self, text, width, bg_color='#E8F0FE',
                 border_color='#4A90D9', font_name='Helvetica',
                 font_size=10):
        Flowable.__init__(self)
        self.text = text
        self.width = width
        self.bg_color = HexColor(bg_color)
        self.border_color = HexColor(border_color)
        self.font_name = font_name
        self.font_size = font_size
        self.padding = 10

    def draw(self):
        # 背景
        self.canv.setFillColor(self.bg_color)
        self.canv.setStrokeColor(self.border_color)
        self.canv.setLineWidth(1)
        self.canv.roundRect(0, 0, self.width,
                           self.height, 5, fill=1, stroke=1)
        # 左邊強調條
        self.canv.setFillColor(self.border_color)
        self.canv.rect(0, 0, 4, self.height, fill=1, stroke=0)
        # 文字
        self.canv.setFillColor(HexColor('#1B2A4A'))
        self.canv.setFont(self.font_name, self.font_size)
        text_obj = self.canv.beginText(self.padding + 6, self.height - self.padding - self.font_size)
        for line in self.text.split('\n'):
            text_obj.textLine(line)
        self.canv.drawText(text_obj)

    def wrap(self, availWidth, availHeight):
        self.width = min(self.width, availWidth)
        lines = self.text.split('\n')
        self.height = len(lines) * (self.font_size + 4) + 2 * self.padding
        return (self.width, self.height)

# 使用
story.append(HighlightBox(
    "投資建議：買入\n目標價：$150\n潛在漲幅：25%",
    width=400
))
```

### 3.3 頁面模板與文檔結構

#### PageTemplate 與 Frame

```python
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

PAGE_WIDTH, PAGE_HEIGHT = A4

class FinancialReportTemplate(BaseDocTemplate):
    """金融報告文檔模板"""

    def __init__(self, filename, **kwargs):
        BaseDocTemplate.__init__(self, filename, pagesize=A4, **kwargs)
        self.page_width = PAGE_WIDTH
        self.page_height = PAGE_HEIGHT
        self._setup_templates()

    def _setup_templates(self):
        """設定不同頁面模板"""

        # 封面模板（無頁首頁尾）
        cover_frame = Frame(
            0, 0, self.page_width, self.page_height,
            id='cover_frame',
            leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
        )
        cover_template = PageTemplate(
            id='Cover',
            frames=[cover_frame],
            onPage=self._draw_cover_background,
        )

        # 標準單欄模板
        main_frame = Frame(
            2.5*cm, 2.5*cm,
            self.page_width - 5*cm,
            self.page_height - 5.5*cm,
            id='main_frame',
        )
        standard_template = PageTemplate(
            id='Standard',
            frames=[main_frame],
            onPage=self._draw_header_footer,
        )

        # 雙欄模板
        left_frame = Frame(
            2.5*cm, 2.5*cm,
            (self.page_width - 6*cm) / 2,
            self.page_height - 5.5*cm,
            id='left_frame',
        )
        right_frame = Frame(
            2.5*cm + (self.page_width - 6*cm) / 2 + 1*cm, 2.5*cm,
            (self.page_width - 6*cm) / 2,
            self.page_height - 5.5*cm,
            id='right_frame',
        )
        two_col_template = PageTemplate(
            id='TwoColumn',
            frames=[left_frame, right_frame],
            onPage=self._draw_header_footer,
        )

        self.addPageTemplates([
            cover_template,
            standard_template,
            two_col_template,
        ])

    def _draw_cover_background(self, canvas, doc):
        """封面背景繪製"""
        canvas.saveState()
        # 漸變背景效果（深藍到中藍）
        for i in range(int(self.page_height)):
            ratio = i / self.page_height
            r = int(27 + (45 - 27) * ratio)
            g = int(42 + (95 - 42) * ratio)
            b = int(74 + (138 - 74) * ratio)
            canvas.setStrokeColorRGB(r/255, g/255, b/255)
            canvas.setLineWidth(1)
            canvas.line(0, i, self.page_width, i)
        canvas.restoreState()

    def _draw_header_footer(self, canvas, doc):
        """標準頁首頁尾繪製"""
        canvas.saveState()

        # ===== 頁首 =====
        # Logo（如有）
        # canvas.drawImage("logo.png", 2.5*cm, self.page_height - 1.8*cm,
        #                  width=2.5*cm, height=1*cm, preserveAspectRatio=True)

        # 報告類別
        canvas.setFont('Helvetica-Bold', 9)
        canvas.setFillColor(HexColor('#1B2A4A'))
        canvas.drawString(2.5*cm, self.page_height - 1.5*cm, "股票研究報告")

        # 日期
        canvas.setFont('Helvetica', 9)
        canvas.drawRightString(
            self.page_width - 2.5*cm,
            self.page_height - 1.5*cm,
            "2026年3月28日"
        )

        # 頁首分隔線
        canvas.setStrokeColor(HexColor('#1B2A4A'))
        canvas.setLineWidth(1)
        canvas.line(2.5*cm, self.page_height - 2*cm,
                   self.page_width - 2.5*cm, self.page_height - 2*cm)

        # ===== 頁尾 =====
        # 頁尾分隔線
        canvas.setStrokeColor(HexColor('#D0D5DD'))
        canvas.setLineWidth(0.5)
        canvas.line(2.5*cm, 2*cm, self.page_width - 2.5*cm, 2*cm)

        # 免責聲明
        canvas.setFont('Helvetica', 6)
        canvas.setFillColor(HexColor('#9E9E9E'))
        canvas.drawString(2.5*cm, 1.5*cm,
            "本報告僅供參考，不構成投資建議。投資有風險，入市需謹慎。")

        # 頁碼
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(HexColor('#616161'))
        canvas.drawRightString(
            self.page_width - 2.5*cm, 1.5*cm,
            f"第 {doc.page} 頁"
        )

        canvas.restoreState()
```

#### 切換頁面模板

```python
from reportlab.platypus import NextPageTemplate, PageBreak

story = []

# 封面（使用 Cover 模板）
story.append(NextPageTemplate('Cover'))
# ...封面內容...

# 切換到標準模板
story.append(NextPageTemplate('Standard'))
story.append(PageBreak())
# ...正文內容...

# 切換到雙欄模板
story.append(NextPageTemplate('TwoColumn'))
story.append(PageBreak())
# ...雙欄內容...
```

#### 動態目錄生成

```python
from reportlab.platypus import TableOfContents
from reportlab.lib.styles import ParagraphStyle

# 目錄樣式
toc = TableOfContents()
toc.levelStyles = [
    ParagraphStyle(
        name='TOC1',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        leftIndent=20,
        firstLineIndent=-20,
        spaceBefore=5,
        spaceAfter=3,
    ),
    ParagraphStyle(
        name='TOC2',
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        leftIndent=40,
        firstLineIndent=-20,
        spaceBefore=2,
        spaceAfter=2,
    ),
    ParagraphStyle(
        name='TOC3',
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        leftIndent=60,
        firstLineIndent=-20,
        spaceBefore=1,
        spaceAfter=1,
    ),
]

# 加入目錄到 Story
story.append(Paragraph("目錄", styles['Heading1']))
story.append(toc)
story.append(PageBreak())

# 使用帶書籤的標題（會自動填入目錄）
# 注意：需使用 multiBuild 而非 build
story.append(Paragraph("第一章：市場概覽", styles['Heading1']))
# ... 內容 ...

# 建構文檔（必須用 multiBuild 才能生成目錄）
doc = FinancialReportTemplate("report.pdf")
doc.multiBuild(story)
```

#### 為標題加入書籤（搭配目錄）

```python
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

class BookmarkParagraph(Paragraph):
    """帶書籤功能的段落，自動加入目錄"""

    def __init__(self, text, style, level=0, bookmark_name=None):
        self.level = level
        self.bookmark_name = bookmark_name or text.replace(' ', '_')
        # 加入書籤錨點的 HTML 標記
        bookmarked_text = f'<a name="{self.bookmark_name}"/>{text}'
        Paragraph.__init__(self, bookmarked_text, style)

    def draw(self):
        # 通知目錄
        key = self.bookmark_name
        self.canv.bookmarkPage(key)
        self.canv.addOutlineEntry(
            self.getPlainText(), key, level=self.level
        )
        Paragraph.draw(self)
```

### 3.4 ReportLab 原生圖表

```python
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.charts.piecharts import Pie
from reportlab.lib.colors import HexColor

# ========== 柱狀圖 ==========
def create_bar_chart(data, categories, title="營收趨勢"):
    """建立 ReportLab 原生柱狀圖"""
    drawing = Drawing(400, 250)

    chart = VerticalBarChart()
    chart.x = 50
    chart.y = 30
    chart.width = 300
    chart.height = 180
    chart.data = data  # [[100, 120, 140, 160]]
    chart.categoryAxis.categoryNames = categories  # ['Q1', 'Q2', 'Q3', 'Q4']

    # 樣式
    chart.bars[0].fillColor = HexColor('#2D5F8A')
    chart.valueAxis.valueMin = 0
    chart.valueAxis.labels.fontName = 'Helvetica'
    chart.valueAxis.labels.fontSize = 8
    chart.categoryAxis.labels.fontName = 'Helvetica'
    chart.categoryAxis.labels.fontSize = 8

    # 標題
    drawing.add(String(200, 230, title,
                       fontName='Helvetica-Bold', fontSize=12,
                       textAnchor='middle',
                       fillColor=HexColor('#1B2A4A')))
    drawing.add(chart)
    return drawing

# ========== 折線圖 ==========
def create_line_chart(data_points, title="股價走勢"):
    """建立 ReportLab 原生折線圖"""
    drawing = Drawing(400, 250)

    chart = LinePlot()
    chart.x = 50
    chart.y = 30
    chart.width = 300
    chart.height = 180
    chart.data = data_points  # [[(0, 100), (1, 105), (2, 98), ...]]

    chart.lines[0].strokeColor = HexColor('#2D5F8A')
    chart.lines[0].strokeWidth = 2

    drawing.add(chart)
    return drawing

# ========== 圓餅圖 ==========
def create_pie_chart(data, labels, title="資產配置"):
    """建立 ReportLab 原生圓餅圖"""
    drawing = Drawing(300, 250)

    pie = Pie()
    pie.x = 75
    pie.y = 30
    pie.width = 150
    pie.height = 150
    pie.data = data  # [40, 30, 20, 10]
    pie.labels = labels  # ['股票', '債券', '黃金', '現金']

    colors = [HexColor(c) for c in ['#2D5F8A', '#C4A35A', '#2E7D32', '#9E9E9E']]
    for i, color in enumerate(colors[:len(data)]):
        pie.slices[i].fillColor = color
        pie.slices[i].strokeColor = HexColor('#FFFFFF')
        pie.slices[i].strokeWidth = 1

    drawing.add(pie)
    return drawing

# 加入 Story
story.append(create_bar_chart([[100, 120, 140, 160]], ['Q1', 'Q2', 'Q3', 'Q4']))
```

#### 嵌入 matplotlib 圖表

```python
import matplotlib.pyplot as plt
import matplotlib
import io
from reportlab.platypus import Image
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import cm

# 設定 matplotlib
matplotlib.rcParams['figure.dpi'] = 150
matplotlib.rcParams['savefig.dpi'] = 300
matplotlib.rcParams['pdf.fonttype'] = 42

def matplotlib_to_reportlab(fig, width=15*cm, height=9*cm):
    """將 matplotlib 圖表轉為 ReportLab Image Flowable"""
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=300,
                bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close(fig)
    buffer.seek(0)
    return Image(ImageReader(buffer), width=width, height=height)

# 範例：建立股價走勢圖
def create_stock_price_chart(dates, prices, title="股價走勢"):
    fig, ax = plt.subplots(figsize=(10, 6))

    ax.plot(dates, prices, color='#2D5F8A', linewidth=2)
    ax.fill_between(dates, prices, alpha=0.1, color='#2D5F8A')

    # MA 均線
    import numpy as np
    if len(prices) >= 20:
        ma20 = np.convolve(prices, np.ones(20)/20, mode='valid')
        ax.plot(dates[19:], ma20, color='#C4A35A', linewidth=1,
                linestyle='--', label='MA20')

    ax.set_title(title, fontsize=14, fontweight='bold', color='#1B2A4A')
    ax.set_ylabel('價格', fontsize=10)
    ax.legend(loc='upper left')
    ax.grid(True, alpha=0.3)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    return fig

# 使用
fig = create_stock_price_chart(dates, prices, "台積電 (2330) 股價走勢")
story.append(matplotlib_to_reportlab(fig))
```

---

## 四、WeasyPrint（HTML → PDF）教程

### 4.1 基礎篇

#### 安裝配置

```bash
# macOS
brew install pango gdk-pixbuf libffi
pip install weasyprint jinja2

# 驗證安裝
python -c "import weasyprint; print(weasyprint.__version__)"
```

#### 最簡範例

```python
from weasyprint import HTML

html_string = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 2cm; }
        h1 { color: #1B2A4A; border-bottom: 2px solid #1B2A4A; }
        table { width: 100%; border-collapse: collapse; }
        th { background: #1B2A4A; color: white; padding: 8px; }
        td { padding: 8px; border-bottom: 1px solid #D0D5DD; }
        tr:nth-child(even) { background: #F5F7FA; }
    </style>
</head>
<body>
    <h1>股票研究報告</h1>
    <p>報告日期：2026年3月28日</p>
    <table>
        <tr><th>指標</th><th>數值</th><th>變化</th></tr>
        <tr><td>營收</td><td>1,250億</td><td style="color:#2E7D32;">+15.9%</td></tr>
        <tr><td>淨利</td><td>180億</td><td style="color:#2E7D32;">+20.9%</td></tr>
    </table>
</body>
</html>
"""

HTML(string=html_string).write_pdf("report.pdf")
```

#### 中文字體配置

```python
# 方法一：使用系統字體（推薦）
# WeasyPrint 自動讀取 Fontconfig 設定的系統字體
css_with_chinese = """
body {
    font-family: 'Microsoft JhengHei', 'Noto Sans TC',
                 'PingFang TC', 'Heiti TC', sans-serif;
}
"""

# 方法二：使用 @font-face 載入字體
css_with_font_face = """
@font-face {
    font-family: 'NotoSansTC';
    src: url('file:///path/to/NotoSansTC-Regular.otf');
    font-weight: normal;
}
@font-face {
    font-family: 'NotoSansTC';
    src: url('file:///path/to/NotoSansTC-Bold.otf');
    font-weight: bold;
}
body {
    font-family: 'NotoSansTC', sans-serif;
}
"""
```

### 4.2 CSS @page 規則（分頁與頁面控制）

```css
/* ========== 頁面基礎設定 ========== */
@page {
    size: A4 portrait;        /* A4 直向 */
    /* size: A4 landscape; */ /* A4 橫向 */
    margin: 2.5cm 2cm 2.5cm 2.5cm;  /* 上 右 下 左 */
}

/* ========== 頁首 ========== */
@page {
    @top-left {
        content: "股票研究報告";
        font-family: 'Microsoft JhengHei', sans-serif;
        font-size: 9pt;
        font-weight: bold;
        color: #1B2A4A;
        border-bottom: 1pt solid #1B2A4A;
        padding-bottom: 5pt;
    }
    @top-right {
        content: "2026年3月28日";
        font-size: 9pt;
        color: #616161;
        border-bottom: 1pt solid #1B2A4A;
        padding-bottom: 5pt;
    }
}

/* ========== 頁尾 ========== */
@page {
    @bottom-center {
        content: "第 " counter(page) " 頁，共 " counter(pages) " 頁";
        font-size: 8pt;
        color: #9E9E9E;
    }
    @bottom-left {
        content: "本報告僅供參考，不構成投資建議";
        font-size: 6pt;
        color: #9E9E9E;
    }
}

/* ========== 首頁不同樣式 ========== */
@page :first {
    margin-top: 0;
    @top-left { content: none; }
    @top-right { content: none; }
    @bottom-center { content: none; }
    @bottom-left { content: none; }
}

/* ========== 命名頁面 ========== */
@page cover {
    margin: 0;
    @top-left { content: none; }
    @bottom-center { content: none; }
}

@page landscape-page {
    size: A4 landscape;
}

/* 使用命名頁面 */
.cover-page { page: cover; }
.landscape-section { page: landscape-page; }
```

#### 分頁控制

```css
/* 強制分頁 */
.page-break {
    page-break-before: always;
}

/* 避免標題與內容分離 */
h1, h2, h3 {
    page-break-after: avoid;
}

/* 表格盡量不分頁 */
table {
    page-break-inside: avoid;
}

/* 圖表區塊不分頁 */
.chart-container {
    page-break-inside: avoid;
}

/* 段落至少保留2行在頁面底部 */
p {
    orphans: 2;
    widows: 2;
}
```

### 4.3 Jinja2 模板引擎

#### 基本模板結構

```
templates/
├── base.html          # 基礎模板
├── cover.html         # 封面
├── stock_report.html  # 股票研究報告模板
├── macro_daily.html   # 宏觀日報模板
└── styles/
    └── report.css     # 共用樣式
```

#### 基礎模板 (base.html)

```html
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="utf-8">
    <style>
        {% block styles %}
        @page {
            size: A4 portrait;
            margin: 2.5cm 2cm 2.5cm 2.5cm;
            @top-left {
                content: "{{ report_type | default('研究報告') }}";
                font-size: 9pt;
                font-weight: bold;
                color: #1B2A4A;
                border-bottom: 1pt solid #1B2A4A;
            }
            @top-right {
                content: "{{ report_date }}";
                font-size: 9pt;
                color: #616161;
                border-bottom: 1pt solid #1B2A4A;
            }
            @bottom-center {
                content: "第 " counter(page) " 頁";
                font-size: 8pt;
                color: #9E9E9E;
            }
        }
        @page :first {
            @top-left { content: none; }
            @top-right { content: none; }
        }

        body {
            font-family: 'Microsoft JhengHei', 'Noto Sans TC', sans-serif;
            font-size: 10pt;
            line-height: 1.4;
            color: #212121;
        }
        h1 {
            color: #1B2A4A;
            font-size: 18pt;
            border-bottom: 2pt solid #1B2A4A;
            padding-bottom: 5pt;
            page-break-after: avoid;
        }
        h2 {
            color: #2D5F8A;
            font-size: 14pt;
            page-break-after: avoid;
        }
        h3 {
            color: #4A90D9;
            font-size: 12pt;
            page-break-after: avoid;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10pt 0;
            page-break-inside: avoid;
        }
        th {
            background: #1B2A4A;
            color: white;
            padding: 6pt 8pt;
            font-size: 9pt;
            text-align: center;
        }
        td {
            padding: 5pt 8pt;
            border-bottom: 0.5pt solid #D0D5DD;
            font-size: 8pt;
        }
        tr:nth-child(even) { background: #F5F7FA; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .text-up { color: #2E7D32; }
        .text-down { color: #C62828; }
        .highlight-box {
            background: #E8F0FE;
            border-left: 4pt solid #4A90D9;
            padding: 10pt 15pt;
            margin: 10pt 0;
            page-break-inside: avoid;
        }
        .chart-img {
            width: 100%;
            max-height: 300pt;
            object-fit: contain;
            page-break-inside: avoid;
        }
        {% endblock %}
    </style>
</head>
<body>
    {% block content %}{% endblock %}
</body>
</html>
```

#### 股票研究報告模板 (stock_report.html)

```html
{% extends "base.html" %}

{% block content %}
<!-- 封面 -->
<div class="cover-page" style="text-align: center; padding-top: 150pt;">
    <div style="background: linear-gradient(135deg, #1B2A4A, #2D5F8A);
                color: white; padding: 40pt; border-radius: 10pt;
                display: inline-block; width: 80%;">
        <h1 style="color: white; border: none; font-size: 24pt;">
            {{ company_name }}（{{ ticker }}）
        </h1>
        <h2 style="color: #C4A35A; font-size: 16pt;">
            投資評級：{{ rating }}
        </h2>
        <div style="display: flex; justify-content: space-around; margin-top: 20pt;">
            <div>
                <div style="font-size: 10pt; opacity: 0.8;">目標價</div>
                <div style="font-size: 20pt; font-weight: bold;">
                    ${{ target_price }}
                </div>
            </div>
            <div>
                <div style="font-size: 10pt; opacity: 0.8;">現價</div>
                <div style="font-size: 20pt; font-weight: bold;">
                    ${{ current_price }}
                </div>
            </div>
            <div>
                <div style="font-size: 10pt; opacity: 0.8;">潛在漲幅</div>
                <div style="font-size: 20pt; font-weight: bold;
                            color: {% if upside > 0 %}#2E7D32{% else %}#C62828{% endif %};">
                    {{ upside }}%
                </div>
            </div>
        </div>
    </div>
    <div style="margin-top: 30pt; color: #616161; font-size: 9pt;">
        <p>分析師：{{ analyst }}</p>
        <p>報告日期：{{ report_date }}</p>
    </div>
</div>

<!-- 投資摘要 -->
<div style="page-break-before: always;">
    <h1>投資摘要</h1>
    <div class="highlight-box">
        {{ investment_summary }}
    </div>

    <h2>核心觀點</h2>
    <ul>
        {% for point in key_points %}
        <li>{{ point }}</li>
        {% endfor %}
    </ul>
</div>

<!-- 財務數據 -->
<div style="page-break-before: always;">
    <h1>財務數據</h1>

    <h2>損益表摘要</h2>
    <table>
        <thead>
            <tr>
                <th>指標</th>
                {% for year in financial_years %}
                <th>{{ year }}</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in income_data %}
            <tr>
                <td><strong>{{ row.metric }}</strong></td>
                {% for val in row.values %}
                <td class="text-right
                    {% if val.change > 0 %}text-up
                    {% elif val.change < 0 %}text-down
                    {% endif %}">
                    {{ val.display }}
                </td>
                {% endfor %}
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <h2>關鍵比率</h2>
    <table>
        <thead>
            <tr>
                <th>指標</th>
                {% for year in financial_years %}
                <th>{{ year }}</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in ratio_data %}
            <tr>
                <td><strong>{{ row.metric }}</strong></td>
                {% for val in row.values %}
                <td class="text-right">{{ val }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
        </tbody>
    </table>
</div>

<!-- 圖表 -->
{% if charts %}
<div style="page-break-before: always;">
    <h1>圖表分析</h1>
    {% for chart in charts %}
    <div style="page-break-inside: avoid; margin-bottom: 20pt;">
        <h3>{{ chart.title }}</h3>
        <img class="chart-img" src="{{ chart.path }}" alt="{{ chart.title }}">
        {% if chart.note %}
        <p style="font-size: 8pt; color: #9E9E9E; text-align: center;">
            {{ chart.note }}
        </p>
        {% endif %}
    </div>
    {% endfor %}
</div>
{% endif %}

<!-- 風險因素 -->
<div style="page-break-before: always;">
    <h1>風險因素</h1>
    {% for risk in risks %}
    <div style="margin-bottom: 10pt;">
        <h3 style="color: #C62828;">{{ risk.title }}</h3>
        <p>{{ risk.description }}</p>
    </div>
    {% endfor %}
</div>

<!-- 免責聲明 -->
<div style="page-break-before: always;">
    <h1>免責聲明</h1>
    <div style="font-size: 8pt; color: #616161; line-height: 1.6;">
        {{ disclaimer }}
    </div>
</div>
{% endblock %}
```

#### Python 渲染代碼

```python
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os

def generate_stock_report(data, output_path):
    """使用 Jinja2 + WeasyPrint 生成股票研究報告"""

    # 設定 Jinja2 模板引擎
    template_dir = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader(os.path.join(template_dir, 'templates')))
    template = env.get_template('stock_report.html')

    # 渲染 HTML
    html_content = template.render(
        report_type="股票研究報告",
        report_date=data['report_date'],
        company_name=data['company_name'],
        ticker=data['ticker'],
        rating=data['rating'],
        target_price=data['target_price'],
        current_price=data['current_price'],
        upside=data['upside'],
        analyst=data['analyst'],
        investment_summary=data['summary'],
        key_points=data['key_points'],
        financial_years=data['years'],
        income_data=data['income_data'],
        ratio_data=data['ratio_data'],
        charts=data.get('charts', []),
        risks=data.get('risks', []),
        disclaimer=data.get('disclaimer', '本報告僅供參考，不構成投資建議。'),
    )

    # 轉換為 PDF
    HTML(string=html_content,
         base_url=template_dir).write_pdf(output_path)

    print(f"報告已生成：{output_path}")
    return output_path
```

### 4.4 Jinja2 進階技巧

#### 自定義過濾器

```python
from jinja2 import Environment

env = Environment(loader=FileSystemLoader('templates'))

# 千分位格式化
@env.filters.setdefault
def format_number(value, decimals=0):
    """格式化數字為千分位"""
    try:
        return f"{float(value):,.{decimals}f}"
    except (ValueError, TypeError):
        return value

# 漲跌 CSS 類別
@env.filters.setdefault
def change_class(value):
    """根據數值正負返回 CSS 類別"""
    try:
        v = float(str(value).replace('%', '').replace('+', ''))
        if v > 0:
            return 'text-up'
        elif v < 0:
            return 'text-down'
    except:
        pass
    return ''

# 在模板中使用
# {{ revenue | format_number(1) }}
# <td class="{{ change | change_class }}">{{ change }}</td>
```

#### 模板宏（Macro）

```html
{# macros.html #}

{% macro financial_table(title, headers, rows) %}
<div style="page-break-inside: avoid; margin-bottom: 15pt;">
    <h3>{{ title }}</h3>
    <table>
        <thead>
            <tr>
                {% for header in headers %}
                <th>{{ header }}</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in rows %}
            <tr>
                {% for cell in row %}
                <td class="{% if loop.index0 > 0 %}text-right{% endif %}
                           {{ cell | change_class }}">
                    {{ cell }}
                </td>
                {% endfor %}
            </tr>
            {% endfor %}
        </tbody>
    </table>
</div>
{% endmacro %}

{% macro rating_badge(rating) %}
<span style="
    display: inline-block;
    padding: 3pt 10pt;
    border-radius: 3pt;
    font-weight: bold;
    color: white;
    background: {% if rating == '買入' %}#2E7D32
                {% elif rating == '持有' %}#F57F17
                {% elif rating == '賣出' %}#C62828
                {% else %}#616161{% endif %};">
    {{ rating }}
</span>
{% endmacro %}

{# 在主模板中使用 #}
{% from "macros.html" import financial_table, rating_badge %}
{{ rating_badge(rating) }}
{{ financial_table("損益表", headers, income_rows) }}
```

---

## 五、完整代碼範例

### 5.1 ReportLab 完整股票研究報告

```python
"""
完整的 ReportLab 股票研究報告生成器
使用 Platypus 排版引擎 + matplotlib 圖表
"""

import io
import datetime
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
    Table, TableStyle, Image, PageBreak, NextPageTemplate,
    KeepTogether, CondPageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# ===== 配色常數 =====
C_PRIMARY = '#1B2A4A'
C_SECONDARY = '#2D5F8A'
C_ACCENT = '#C4A35A'
C_UP = '#2E7D32'
C_DOWN = '#C62828'
C_GRAY = '#616161'
C_LIGHT_GRAY = '#9E9E9E'
C_BG_LIGHT = '#F5F7FA'
C_BG_HIGHLIGHT = '#E8F0FE'
C_BORDER = '#D0D5DD'
C_WHITE = '#FFFFFF'

# ===== matplotlib 設定 =====
matplotlib.rcParams.update({
    'figure.dpi': 150,
    'savefig.dpi': 300,
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'pdf.fonttype': 42,
    'font.family': 'sans-serif',
    'axes.unicode_minus': False,
})

PAGE_WIDTH, PAGE_HEIGHT = A4


class StockReportGenerator:
    """股票研究報告 PDF 生成器"""

    def __init__(self, output_path):
        self.output_path = output_path
        self.story = []
        self.styles = self._create_styles()
        self.doc = self._create_doc()

    def _create_styles(self):
        """建立所有段落樣式"""
        styles = getSampleStyleSheet()

        # 封面標題
        styles.add(ParagraphStyle(
            name='CoverTitle',
            fontName='Helvetica-Bold',
            fontSize=28,
            leading=34,
            textColor=HexColor(C_WHITE),
            alignment=TA_CENTER,
            spaceAfter=10,
        ))

        # 封面副標題
        styles.add(ParagraphStyle(
            name='CoverSubtitle',
            fontName='Helvetica',
            fontSize=14,
            leading=18,
            textColor=HexColor(C_ACCENT),
            alignment=TA_CENTER,
            spaceAfter=6,
        ))

        # 章節標題 H1
        styles.add(ParagraphStyle(
            name='H1',
            fontName='Helvetica-Bold',
            fontSize=18,
            leading=22,
            textColor=HexColor(C_PRIMARY),
            spaceBefore=20,
            spaceAfter=12,
            borderWidth=0,
        ))

        # 次標題 H2
        styles.add(ParagraphStyle(
            name='H2',
            fontName='Helvetica-Bold',
            fontSize=13,
            leading=16,
            textColor=HexColor(C_SECONDARY),
            spaceBefore=14,
            spaceAfter=8,
        ))

        # 小標題 H3
        styles.add(ParagraphStyle(
            name='H3',
            fontName='Helvetica-Bold',
            fontSize=11,
            leading=14,
            textColor=HexColor(C_GRAY),
            spaceBefore=10,
            spaceAfter=6,
        ))

        # 正文
        styles.add(ParagraphStyle(
            name='Body',
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            textColor=HexColor('#212121'),
            spaceAfter=6,
            alignment=TA_JUSTIFY,
        ))

        # 重點文字
        styles.add(ParagraphStyle(
            name='Highlight',
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            textColor=HexColor(C_PRIMARY),
            backColor=HexColor(C_BG_HIGHLIGHT),
            borderWidth=0,
            borderPadding=10,
            spaceAfter=12,
        ))

        # 腳注
        styles.add(ParagraphStyle(
            name='Footnote',
            fontName='Helvetica',
            fontSize=7,
            leading=9,
            textColor=HexColor(C_LIGHT_GRAY),
            spaceAfter=3,
        ))

        # 表格標題
        styles.add(ParagraphStyle(
            name='TableCaption',
            fontName='Helvetica-Bold',
            fontSize=9,
            leading=11,
            textColor=HexColor(C_GRAY),
            spaceAfter=4,
            spaceBefore=8,
        ))

        return styles

    def _create_doc(self):
        """建立文檔模板"""
        doc = BaseDocTemplate(
            self.output_path,
            pagesize=A4,
            topMargin=2.5*cm,
            bottomMargin=2.5*cm,
            leftMargin=2.5*cm,
            rightMargin=2*cm,
        )

        # 封面框架
        cover_frame = Frame(
            0, 0, PAGE_WIDTH, PAGE_HEIGHT,
            id='cover', leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
        )

        # 內容框架
        content_frame = Frame(
            2.5*cm, 2.5*cm,
            PAGE_WIDTH - 4.5*cm,
            PAGE_HEIGHT - 5.5*cm,
            id='content',
        )

        doc.addPageTemplates([
            PageTemplate(id='Cover', frames=[cover_frame],
                        onPage=self._draw_cover_bg),
            PageTemplate(id='Content', frames=[content_frame],
                        onPage=self._draw_header_footer),
        ])

        return doc

    def _draw_cover_bg(self, canvas, doc):
        """繪製封面背景"""
        canvas.saveState()
        # 漸層背景
        steps = 200
        for i in range(steps):
            ratio = i / steps
            r = (27 + (45 - 27) * ratio) / 255
            g = (42 + (95 - 42) * ratio) / 255
            b = (74 + (138 - 74) * ratio) / 255
            canvas.setFillColorRGB(r, g, b)
            y = PAGE_HEIGHT * (1 - ratio)
            h = PAGE_HEIGHT / steps + 1
            canvas.rect(0, y - h, PAGE_WIDTH, h, fill=1, stroke=0)
        canvas.restoreState()

    def _draw_header_footer(self, canvas, doc):
        """繪製頁首頁尾"""
        canvas.saveState()

        # 頁首
        canvas.setFont('Helvetica-Bold', 9)
        canvas.setFillColor(HexColor(C_PRIMARY))
        canvas.drawString(2.5*cm, PAGE_HEIGHT - 1.5*cm, "Equity Research")

        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(HexColor(C_GRAY))
        canvas.drawRightString(PAGE_WIDTH - 2*cm, PAGE_HEIGHT - 1.5*cm,
                              datetime.date.today().strftime("%Y-%m-%d"))

        canvas.setStrokeColor(HexColor(C_PRIMARY))
        canvas.setLineWidth(1)
        canvas.line(2.5*cm, PAGE_HEIGHT - 2*cm,
                   PAGE_WIDTH - 2*cm, PAGE_HEIGHT - 2*cm)

        # 頁尾
        canvas.setStrokeColor(HexColor(C_BORDER))
        canvas.setLineWidth(0.5)
        canvas.line(2.5*cm, 2*cm, PAGE_WIDTH - 2*cm, 2*cm)

        canvas.setFont('Helvetica', 6)
        canvas.setFillColor(HexColor(C_LIGHT_GRAY))
        canvas.drawString(2.5*cm, 1.5*cm,
            "This report is for informational purposes only. "
            "Not investment advice.")

        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(HexColor(C_GRAY))
        canvas.drawRightString(PAGE_WIDTH - 2*cm, 1.5*cm,
                              f"Page {doc.page}")

        canvas.restoreState()

    def add_cover(self, company, ticker, rating, target_price,
                  current_price, analyst):
        """加入封面"""
        self.story.append(NextPageTemplate('Cover'))
        self.story.append(Spacer(1, 8*cm))

        self.story.append(Paragraph(
            f"{company}", self.styles['CoverTitle']))
        self.story.append(Paragraph(
            f"({ticker})", self.styles['CoverSubtitle']))
        self.story.append(Spacer(1, 1*cm))

        # 評級與價格資訊
        upside = ((target_price - current_price) / current_price) * 100
        info_data = [
            ['Rating', 'Target Price', 'Current Price', 'Upside'],
            [rating, f'${target_price:.2f}',
             f'${current_price:.2f}', f'{upside:.1f}%'],
        ]
        info_table = Table(info_data, colWidths=[100, 100, 100, 80])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#FFFFFF20')),
            ('TEXTCOLOR', (0, 0), (-1, 0), HexColor(C_ACCENT)),
            ('TEXTCOLOR', (0, 1), (-1, 1), HexColor(C_WHITE)),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 16),
            ('ALIGNMENT', (0, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        self.story.append(info_table)
        self.story.append(Spacer(1, 3*cm))

        # 分析師資訊
        self.story.append(Paragraph(
            f'<font color="{C_LIGHT_GRAY}" size="9">'
            f'Analyst: {analyst} | '
            f'{datetime.date.today().strftime("%B %d, %Y")}'
            f'</font>',
            ParagraphStyle('CoverInfo', alignment=TA_CENTER)
        ))

        # 切換到內容模板
        self.story.append(NextPageTemplate('Content'))
        self.story.append(PageBreak())

    def add_section(self, title, content, level=1):
        """加入章節"""
        style_name = f'H{level}'
        if style_name in [s.name for s in self.styles.byName.values()]:
            self.story.append(Paragraph(title, self.styles[style_name]))
        if content:
            self.story.append(Paragraph(content, self.styles['Body']))

    def add_highlight_box(self, text):
        """加入重點提示框"""
        self.story.append(Paragraph(text, self.styles['Highlight']))

    def add_financial_table(self, title, headers, data,
                            col_widths=None, highlight_col=None):
        """加入財務數據表格"""
        self.story.append(Paragraph(title, self.styles['TableCaption']))

        table_data = [headers] + data
        if col_widths is None:
            col_widths = [100] + [70] * (len(headers) - 1)

        table = Table(table_data, colWidths=col_widths)

        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(C_PRIMARY)),
            ('TEXTCOLOR', (0, 0), (-1, 0), HexColor(C_WHITE)),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGNMENT', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGNMENT', (0, 1), (0, -1), 'LEFT'),
            ('ALIGNMENT', (1, 1), (-1, -1), 'RIGHT'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [HexColor(C_WHITE), HexColor(C_BG_LIGHT)]),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor(C_BORDER)),
            ('LINEBELOW', (0, 0), (-1, 0), 1.5, HexColor(C_PRIMARY)),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]

        # 條件格式：漲跌配色
        if highlight_col is not None:
            for row_idx in range(1, len(table_data)):
                val = str(table_data[row_idx][highlight_col])
                try:
                    num = float(val.replace('%', '').replace('+', '').replace(',', ''))
                    if num > 0:
                        style_cmds.append(('TEXTCOLOR',
                            (highlight_col, row_idx),
                            (highlight_col, row_idx),
                            HexColor(C_UP)))
                    elif num < 0:
                        style_cmds.append(('TEXTCOLOR',
                            (highlight_col, row_idx),
                            (highlight_col, row_idx),
                            HexColor(C_DOWN)))
                except (ValueError, TypeError):
                    pass

        table.setStyle(TableStyle(style_cmds))
        self.story.append(KeepTogether([
            Spacer(1, 2*mm),
            table,
            Spacer(1, 4*mm),
        ]))

    def add_matplotlib_chart(self, fig, width=15*cm, height=9*cm):
        """嵌入 matplotlib 圖表"""
        buffer = io.BytesIO()
        fig.savefig(buffer, format='png', dpi=300,
                    bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        buffer.seek(0)
        img = Image(ImageReader(buffer), width=width, height=height)
        self.story.append(KeepTogether([img, Spacer(1, 5*mm)]))

    def add_page_break(self):
        """加入分頁"""
        self.story.append(PageBreak())

    def build(self):
        """建構 PDF"""
        self.doc.build(self.story)
        print(f"報告已生成：{self.output_path}")


# ===== 使用範例 =====
def generate_sample_report():
    """生成範例股票研究報告"""
    report = StockReportGenerator("stock_research_report.pdf")

    # 1. 封面
    report.add_cover(
        company="Taiwan Semiconductor (TSMC)",
        ticker="2330.TW",
        rating="BUY",
        target_price=850.00,
        current_price=680.00,
        analyst="Jamie Chen",
    )

    # 2. 投資摘要
    report.add_section("Investment Summary", None)
    report.add_highlight_box(
        "<b>BUY Rating Maintained</b> - Target Price NT$850<br/><br/>"
        "We maintain our BUY rating on TSMC with a 12-month target "
        "price of NT$850, representing 25% upside. Key catalysts include: "
        "(1) Strong AI chip demand, (2) Advanced packaging momentum, "
        "(3) CoWoS capacity expansion."
    )

    report.add_section("Key Investment Points", None)
    points = [
        "AI-driven demand continues to accelerate HPC revenue growth",
        "3nm/2nm technology leadership provides sustainable competitive moat",
        "CoWoS advanced packaging capacity doubling in 2026",
        "Overseas fab investments strengthen geopolitical positioning",
    ]
    for p in points:
        report.story.append(Paragraph(
            f"<bullet>&bull;</bullet> {p}",
            report.styles['Body']))

    # 3. 財務數據
    report.add_page_break()
    report.add_section("Financial Summary", None)

    report.add_financial_table(
        "Table 1: Income Statement Summary (NT$ Billion)",
        ['Metric', 'FY2023', 'FY2024', 'FY2025E', 'FY2026E', 'YoY%'],
        [
            ['Revenue', '2,162', '2,894', '3,580', '4,250', '+18.7%'],
            ['Gross Profit', '1,153', '1,582', '2,005', '2,422', '+20.8%'],
            ['Operating Income', '934', '1,320', '1,700', '2,080', '+22.4%'],
            ['Net Income', '838', '1,173', '1,510', '1,840', '+21.9%'],
            ['EPS (NT$)', '32.3', '45.3', '58.2', '71.0', '+22.0%'],
        ],
        col_widths=[110, 65, 65, 65, 65, 55],
        highlight_col=5,
    )

    report.add_financial_table(
        "Table 2: Key Ratios",
        ['Metric', 'FY2023', 'FY2024', 'FY2025E', 'FY2026E'],
        [
            ['Gross Margin', '53.3%', '54.7%', '56.0%', '57.0%'],
            ['Operating Margin', '43.2%', '45.6%', '47.5%', '48.9%'],
            ['Net Margin', '38.8%', '40.5%', '42.2%', '43.3%'],
            ['ROE', '28.5%', '32.1%', '35.0%', '37.5%'],
            ['P/E (x)', '21.1x', '15.0x', '11.7x', '9.6x'],
        ],
        col_widths=[110, 65, 65, 65, 65],
    )

    # 4. 圖表
    report.add_page_break()
    report.add_section("Charts & Analysis", None)

    # 營收趨勢圖
    fig, ax = plt.subplots(figsize=(10, 5))
    years = ['2021', '2022', '2023', '2024', '2025E', '2026E']
    revenue = [1587, 2264, 2162, 2894, 3580, 4250]
    bars = ax.bar(years, revenue, color=C_SECONDARY, edgecolor='white', width=0.6)
    # 標註數值
    for bar, val in zip(bars, revenue):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
                f'{val:,}', ha='center', va='bottom', fontsize=9,
                fontweight='bold', color=C_PRIMARY)
    ax.set_title('Revenue Trend (NT$ Billion)', fontsize=14,
                 fontweight='bold', color=C_PRIMARY, pad=15)
    ax.set_ylabel('NT$ Billion', fontsize=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3)
    report.add_matplotlib_chart(fig)

    # 利潤率趨勢圖
    fig, ax = plt.subplots(figsize=(10, 5))
    gross_m = [51.6, 59.6, 53.3, 54.7, 56.0, 57.0]
    op_m = [41.0, 49.5, 43.2, 45.6, 47.5, 48.9]
    net_m = [37.6, 44.4, 38.8, 40.5, 42.2, 43.3]
    ax.plot(years, gross_m, 'o-', color=C_SECONDARY, linewidth=2, label='Gross Margin')
    ax.plot(years, op_m, 's--', color=C_ACCENT, linewidth=2, label='Operating Margin')
    ax.plot(years, net_m, '^:', color=C_UP, linewidth=2, label='Net Margin')
    ax.set_title('Profitability Trend (%)', fontsize=14,
                 fontweight='bold', color=C_PRIMARY, pad=15)
    ax.set_ylabel('%', fontsize=10)
    ax.legend(loc='lower right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, alpha=0.3)
    ax.set_ylim(30, 65)
    report.add_matplotlib_chart(fig)

    # 5. 風險因素
    report.add_page_break()
    report.add_section("Risk Factors", None)
    risks = [
        ("Geopolitical Risk",
         "Escalation of US-China tensions could disrupt supply chains "
         "and impact customer orders."),
        ("Technology Competition",
         "Samsung and Intel are investing heavily in advanced nodes. "
         "Any breakthrough could narrow TSMC's technology lead."),
        ("Cyclical Demand Risk",
         "Semiconductor industry is cyclical. A demand downturn "
         "could impact utilization rates and margins."),
        ("Capex Execution Risk",
         "Massive capital expenditure plans carry execution risk "
         "and could pressure free cash flow in the short term."),
    ]
    for title, desc in risks:
        report.add_section(title, desc, level=3)

    # 6. 免責聲明
    report.add_page_break()
    report.add_section("Disclaimer", None)
    report.story.append(Paragraph(
        "This research report is prepared for informational purposes only "
        "and does not constitute investment advice. The information contained "
        "herein is based on sources believed to be reliable but is not "
        "guaranteed. Past performance is not indicative of future results. "
        "Investors should conduct their own due diligence before making "
        "investment decisions.",
        report.styles['Footnote']
    ))

    # 建構
    report.build()

# 執行
# generate_sample_report()
```

### 5.2 WeasyPrint 宏觀日報完整範例

```python
"""
WeasyPrint + Jinja2 宏觀經濟日報生成器
HTML 模板驅動，支持中文
"""

import os
import datetime
import base64
import io
import matplotlib.pyplot as plt
import matplotlib
from jinja2 import Environment, BaseLoader
from weasyprint import HTML

matplotlib.rcParams.update({
    'figure.dpi': 150,
    'savefig.dpi': 200,
    'figure.facecolor': 'white',
})

# ===== HTML 模板（內嵌，也可改為外部檔案） =====
DAILY_REPORT_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<style>
    @page {
        size: A4 portrait;
        margin: 2cm 1.8cm 2cm 2cm;

        @top-left {
            content: "宏觀經濟日報";
            font-family: 'Microsoft JhengHei', 'Noto Sans TC', sans-serif;
            font-size: 8pt;
            font-weight: bold;
            color: #1B2A4A;
            border-bottom: 0.5pt solid #1B2A4A;
            padding-bottom: 3pt;
        }
        @top-right {
            content: "{{ report_date }}";
            font-size: 8pt;
            color: #616161;
            border-bottom: 0.5pt solid #1B2A4A;
            padding-bottom: 3pt;
        }
        @bottom-center {
            content: "第 " counter(page) " 頁 / 共 " counter(pages) " 頁";
            font-size: 7pt;
            color: #9E9E9E;
        }
    }

    @page :first {
        margin-top: 1cm;
        @top-left { content: none; border: none; }
        @top-right { content: none; border: none; }
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
        font-family: 'Microsoft JhengHei', 'Noto Sans TC',
                     'PingFang TC', Arial, sans-serif;
        font-size: 9pt;
        line-height: 1.4;
        color: #212121;
    }

    .header-bar {
        background: linear-gradient(135deg, #1B2A4A, #2D5F8A);
        color: white;
        padding: 20pt 25pt;
        margin: -1cm -1.8cm 15pt -2cm;
        width: calc(100% + 3.8cm);
    }
    .header-bar h1 {
        font-size: 20pt;
        margin-bottom: 5pt;
    }
    .header-bar .subtitle {
        font-size: 10pt;
        color: #C4A35A;
    }

    h2 {
        color: #1B2A4A;
        font-size: 13pt;
        border-bottom: 2pt solid #1B2A4A;
        padding-bottom: 4pt;
        margin: 15pt 0 8pt 0;
        page-break-after: avoid;
    }

    h3 {
        color: #2D5F8A;
        font-size: 10pt;
        margin: 10pt 0 5pt 0;
        page-break-after: avoid;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        margin: 8pt 0;
        page-break-inside: avoid;
        font-size: 8pt;
    }
    th {
        background: #1B2A4A;
        color: white;
        padding: 5pt 6pt;
        text-align: center;
        font-weight: bold;
    }
    td {
        padding: 4pt 6pt;
        border-bottom: 0.5pt solid #D0D5DD;
    }
    tr:nth-child(even) { background: #F5F7FA; }
    .text-right { text-align: right; }
    .text-center { text-align: center; }
    .up { color: #2E7D32; font-weight: bold; }
    .down { color: #C62828; font-weight: bold; }

    .summary-box {
        background: #E8F0FE;
        border-left: 4pt solid #4A90D9;
        padding: 10pt 15pt;
        margin: 10pt 0;
        page-break-inside: avoid;
    }

    .two-col {
        display: flex;
        gap: 15pt;
    }
    .two-col > div {
        flex: 1;
    }

    .chart-container {
        text-align: center;
        margin: 10pt 0;
        page-break-inside: avoid;
    }
    .chart-container img {
        max-width: 100%;
        max-height: 250pt;
    }
    .chart-caption {
        font-size: 7pt;
        color: #9E9E9E;
        margin-top: 3pt;
    }

    .footer-disclaimer {
        font-size: 6pt;
        color: #9E9E9E;
        margin-top: 20pt;
        border-top: 0.5pt solid #D0D5DD;
        padding-top: 5pt;
    }
</style>
</head>
<body>

<!-- 頁首標題列 -->
<div class="header-bar">
    <h1>宏觀經濟日報</h1>
    <div class="subtitle">{{ report_date }} | 全球市場概覽</div>
</div>

<!-- 今日摘要 -->
<div class="summary-box">
    <strong>今日重點：</strong>{{ daily_summary }}
</div>

<!-- 全球指數 -->
<h2>全球主要指數</h2>
<table>
    <thead>
        <tr>
            <th>指數</th>
            <th>收盤價</th>
            <th>漲跌</th>
            <th>漲跌幅</th>
            <th>本週累計</th>
            <th>年初至今</th>
        </tr>
    </thead>
    <tbody>
        {% for idx in indices %}
        <tr>
            <td><strong>{{ idx.name }}</strong></td>
            <td class="text-right">{{ idx.close }}</td>
            <td class="text-right {{ 'up' if idx.change > 0 else 'down' if idx.change < 0 else '' }}">
                {{ '%+.2f' | format(idx.change) }}
            </td>
            <td class="text-right {{ 'up' if idx.pct > 0 else 'down' if idx.pct < 0 else '' }}">
                {{ '%+.2f%%' | format(idx.pct) }}
            </td>
            <td class="text-right {{ 'up' if idx.wtd > 0 else 'down' if idx.wtd < 0 else '' }}">
                {{ '%+.2f%%' | format(idx.wtd) }}
            </td>
            <td class="text-right {{ 'up' if idx.ytd > 0 else 'down' if idx.ytd < 0 else '' }}">
                {{ '%+.2f%%' | format(idx.ytd) }}
            </td>
        </tr>
        {% endfor %}
    </tbody>
</table>

<!-- 圖表區 -->
{% if charts %}
<h2>市場走勢圖</h2>
<div class="two-col">
    {% for chart in charts %}
    <div class="chart-container">
        <img src="data:image/png;base64,{{ chart.base64 }}" alt="{{ chart.title }}">
        <div class="chart-caption">{{ chart.title }}</div>
    </div>
    {% endfor %}
</div>
{% endif %}

<!-- 經濟數據 -->
{% if economic_data %}
<h2 style="page-break-before: always;">經濟數據追蹤</h2>
<table>
    <thead>
        <tr>
            <th>指標</th>
            <th>最新值</th>
            <th>前期值</th>
            <th>預期值</th>
            <th>發布日期</th>
        </tr>
    </thead>
    <tbody>
        {% for item in economic_data %}
        <tr>
            <td><strong>{{ item.name }}</strong></td>
            <td class="text-right">{{ item.actual }}</td>
            <td class="text-right">{{ item.previous }}</td>
            <td class="text-right">{{ item.expected }}</td>
            <td class="text-center">{{ item.date }}</td>
        </tr>
        {% endfor %}
    </tbody>
</table>
{% endif %}

<!-- 央行政策 -->
{% if central_bank_notes %}
<h2>央行政策動態</h2>
{% for note in central_bank_notes %}
<h3>{{ note.bank }}</h3>
<p>{{ note.content }}</p>
{% endfor %}
{% endif %}

<!-- 免責聲明 -->
<div class="footer-disclaimer">
    本報告僅供參考，不構成任何投資建議。報告中的資訊來源被認為是可靠的，
    但不保證其準確性或完整性。投資有風險，入市需謹慎。
    過往業績不代表未來表現。
</div>

</body>
</html>
"""


def fig_to_base64(fig):
    """將 matplotlib 圖表轉為 Base64 字串（供 HTML img 標籤使用）"""
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=200,
                bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')


def generate_macro_daily(data, output_path):
    """
    生成宏觀經濟日報 PDF

    參數：
        data: dict，包含以下鍵：
            - report_date: str
            - daily_summary: str
            - indices: list of dict
            - charts: list of dict (含 base64 圖片)
            - economic_data: list of dict (可選)
            - central_bank_notes: list of dict (可選)
        output_path: str，PDF 輸出路徑
    """
    env = Environment(loader=BaseLoader())
    template = env.from_string(DAILY_REPORT_TEMPLATE)

    html_content = template.render(**data)

    HTML(string=html_content).write_pdf(output_path)
    print(f"宏觀日報已生成：{output_path}")
    return output_path


# ===== 使用範例 =====
def generate_sample_daily():
    """生成範例宏觀日報"""

    # 建立範例圖表
    charts = []

    # 圖表1：S&P 500 走勢
    fig, ax = plt.subplots(figsize=(5, 3))
    import numpy as np
    np.random.seed(42)
    days = range(60)
    prices = 5200 + np.cumsum(np.random.randn(60) * 20)
    ax.plot(days, prices, color='#2D5F8A', linewidth=1.5)
    ax.fill_between(days, prices, alpha=0.1, color='#2D5F8A')
    ax.set_title('S&P 500 (60D)', fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    charts.append({'title': 'S&P 500 近60日走勢', 'base64': fig_to_base64(fig)})

    # 圖表2：美元指數
    fig, ax = plt.subplots(figsize=(5, 3))
    dxy = 104 + np.cumsum(np.random.randn(60) * 0.3)
    ax.plot(days, dxy, color='#C4A35A', linewidth=1.5)
    ax.set_title('DXY (60D)', fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    charts.append({'title': '美元指數近60日走勢', 'base64': fig_to_base64(fig)})

    data = {
        'report_date': datetime.date.today().strftime('%Y年%m月%d日'),
        'daily_summary': (
            '美股三大指數收漲，S&P 500 上漲0.8%，受AI概念股帶動；'
            '美國10年期公債殖利率小幅回落至4.25%；'
            '美元指數持平於104.2。'
        ),
        'indices': [
            {'name': 'S&P 500', 'close': '5,328.42',
             'change': 42.15, 'pct': 0.80, 'wtd': 1.25, 'ytd': 8.50},
            {'name': 'NASDAQ', 'close': '16,892.30',
             'change': 158.20, 'pct': 0.95, 'wtd': 1.80, 'ytd': 12.30},
            {'name': 'Dow Jones', 'close': '39,150.80',
             'change': 120.50, 'pct': 0.31, 'wtd': 0.65, 'ytd': 5.20},
            {'name': 'Nikkei 225', 'close': '40,250.00',
             'change': -85.30, 'pct': -0.21, 'wtd': 0.40, 'ytd': 15.60},
            {'name': 'Hang Seng', 'close': '17,850.20',
             'change': 125.80, 'pct': 0.71, 'wtd': -0.30, 'ytd': 2.10},
            {'name': 'TAIEX', 'close': '20,580.50',
             'change': 165.30, 'pct': 0.81, 'wtd': 1.50, 'ytd': 10.80},
        ],
        'charts': charts,
        'economic_data': [
            {'name': 'US GDP (Q4)', 'actual': '3.2%', 'previous': '2.8%',
             'expected': '3.0%', 'date': '2026-03-25'},
            {'name': 'US CPI (Feb)', 'actual': '2.8%', 'previous': '3.0%',
             'expected': '2.9%', 'date': '2026-03-12'},
            {'name': 'US NFP (Feb)', 'actual': '275K', 'previous': '229K',
             'expected': '200K', 'date': '2026-03-07'},
            {'name': 'China PMI (Feb)', 'actual': '50.2', 'previous': '49.8',
             'expected': '50.0', 'date': '2026-03-01'},
        ],
        'central_bank_notes': [
            {'bank': 'Federal Reserve (Fed)',
             'content': '聯準會維持利率不變於5.25%-5.50%，鮑威爾暗示今年可能降息兩次，'
                        '但強調需看到更多通膨降溫數據。點陣圖中位數指向年底前降息50bp。'},
            {'bank': 'European Central Bank (ECB)',
             'content': '歐央行維持利率不變，拉加德表示通膨正朝目標方向前進，'
                        '市場預期6月可能開始降息。'},
        ],
    }

    generate_macro_daily(data, "macro_daily_report.pdf")

# 執行
# generate_sample_daily()
```

### 5.3 批量生成多支股票 PDF 報告

```python
"""
批量股票 PDF 報告生成器
支持從 CSV/API 讀取數據，批量生成多支股票報告
"""

import os
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

# 設定日誌
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class BatchReportGenerator:
    """批量報告生成器"""

    def __init__(self, output_dir, max_workers=4):
        self.output_dir = output_dir
        self.max_workers = max_workers
        os.makedirs(output_dir, exist_ok=True)

    def generate_single_report(self, stock_data):
        """生成單支股票報告"""
        ticker = stock_data['ticker']
        try:
            output_path = os.path.join(
                self.output_dir,
                f"{ticker}_research_{stock_data.get('date', 'latest')}.pdf"
            )

            # 使用前面定義的 StockReportGenerator
            report = StockReportGenerator(output_path)
            report.add_cover(
                company=stock_data['company_name'],
                ticker=ticker,
                rating=stock_data['rating'],
                target_price=stock_data['target_price'],
                current_price=stock_data['current_price'],
                analyst=stock_data.get('analyst', 'Research Team'),
            )

            # 加入財務數據
            if 'financials' in stock_data:
                report.add_section("Financial Summary", None)
                report.add_financial_table(
                    "Income Statement",
                    stock_data['financials']['headers'],
                    stock_data['financials']['data'],
                )

            report.build()
            logger.info(f"已完成：{ticker} → {output_path}")
            return {'ticker': ticker, 'status': 'success', 'path': output_path}

        except Exception as e:
            logger.error(f"失敗：{ticker} — {str(e)}")
            return {'ticker': ticker, 'status': 'error', 'error': str(e)}

    def generate_batch(self, stock_list):
        """
        批量生成報告

        參數：
            stock_list: list of dict，每個 dict 包含股票數據
        """
        results = []
        start_time = time.time()
        logger.info(f"開始批量生成 {len(stock_list)} 份報告...")

        # 使用執行緒池並行生成
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(self.generate_single_report, data): data['ticker']
                for data in stock_list
            }

            for future in as_completed(futures):
                ticker = futures[future]
                try:
                    result = future.result()
                    results.append(result)
                except Exception as e:
                    results.append({
                        'ticker': ticker,
                        'status': 'error',
                        'error': str(e),
                    })

        elapsed = time.time() - start_time
        success = sum(1 for r in results if r['status'] == 'success')
        failed = sum(1 for r in results if r['status'] == 'error')

        logger.info(
            f"批量生成完成：成功 {success}，失敗 {failed}，"
            f"耗時 {elapsed:.1f} 秒"
        )

        return results

    def generate_from_csv(self, csv_path):
        """從 CSV 檔案讀取股票清單並生成報告"""
        import pandas as pd

        df = pd.read_csv(csv_path)
        stock_list = df.to_dict('records')
        return self.generate_batch(stock_list)


# ===== 使用範例 =====
def batch_example():
    """批量生成範例"""
    stocks = [
        {
            'ticker': '2330.TW',
            'company_name': 'TSMC',
            'rating': 'BUY',
            'target_price': 850,
            'current_price': 680,
            'date': '2026-03-28',
        },
        {
            'ticker': '2317.TW',
            'company_name': 'Hon Hai (Foxconn)',
            'rating': 'HOLD',
            'target_price': 200,
            'current_price': 185,
            'date': '2026-03-28',
        },
        {
            'ticker': 'AAPL',
            'company_name': 'Apple Inc.',
            'rating': 'BUY',
            'target_price': 220,
            'current_price': 195,
            'date': '2026-03-28',
        },
    ]

    generator = BatchReportGenerator(
        output_dir="./reports/batch_output",
        max_workers=3,
    )
    results = generator.generate_batch(stocks)

    # 輸出結果摘要
    for r in results:
        status = "OK" if r['status'] == 'success' else f"FAIL: {r.get('error')}"
        print(f"  {r['ticker']}: {status}")

# 執行
# batch_example()
```

---

## 六、自動化報告流程

### 6.1 數據 → 圖表 → PDF 全自動流程

```python
"""
完整的自動化報告管線（Pipeline）
數據獲取 → 分析 → 圖表 → PDF → 通知
"""

import os
import datetime
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

logger = logging.getLogger(__name__)


class ReportPipeline:
    """自動化報告管線"""

    def __init__(self, config):
        """
        config = {
            'output_dir': './reports',
            'data_source': 'yfinance',  # 或 'api', 'csv'
            'report_type': 'daily',     # 'daily', 'weekly', 'stock'
            'email': {                  # 可選
                'smtp_server': 'smtp.gmail.com',
                'smtp_port': 587,
                'sender': 'xxx@gmail.com',
                'password': 'app_password',
                'recipients': ['user@example.com'],
            }
        }
        """
        self.config = config
        self.output_dir = config.get('output_dir', './reports')
        os.makedirs(self.output_dir, exist_ok=True)

    def step1_fetch_data(self):
        """步驟一：獲取數據"""
        logger.info("步驟 1/5：獲取市場數據...")

        data = {}

        # 使用 yfinance 獲取數據
        try:
            import yfinance as yf

            # 主要指數
            indices = {
                'S&P 500': '^GSPC',
                'NASDAQ': '^IXIC',
                'Dow Jones': '^DJI',
                'Nikkei 225': '^N225',
                'Hang Seng': '^HSI',
            }

            index_data = []
            for name, symbol in indices.items():
                ticker = yf.Ticker(symbol)
                hist = ticker.history(period='5d')
                if not hist.empty:
                    close = hist['Close'].iloc[-1]
                    prev = hist['Close'].iloc[-2] if len(hist) > 1 else close
                    change = close - prev
                    pct = (change / prev) * 100
                    index_data.append({
                        'name': name,
                        'close': f'{close:,.2f}',
                        'change': round(change, 2),
                        'pct': round(pct, 2),
                        'wtd': 0,  # 需要更多歷史數據計算
                        'ytd': 0,
                    })

            data['indices'] = index_data
            logger.info(f"  已獲取 {len(index_data)} 個指數數據")

        except ImportError:
            logger.warning("  yfinance 未安裝，使用模擬數據")
            data['indices'] = []  # 使用模擬數據

        return data

    def step2_analyze(self, raw_data):
        """步驟二：數據分析"""
        logger.info("步驟 2/5：數據分析...")

        analysis = {
            'report_date': datetime.date.today().strftime('%Y年%m月%d日'),
            'daily_summary': self._generate_summary(raw_data),
        }
        analysis.update(raw_data)
        return analysis

    def step3_create_charts(self, analysis):
        """步驟三：生成圖表"""
        logger.info("步驟 3/5：生成圖表...")
        import matplotlib.pyplot as plt
        import io
        import base64

        charts = []

        # 指數漲跌幅對比柱狀圖
        if analysis.get('indices'):
            fig, ax = plt.subplots(figsize=(6, 3.5))
            names = [idx['name'] for idx in analysis['indices']]
            pcts = [idx['pct'] for idx in analysis['indices']]
            colors = ['#2E7D32' if p >= 0 else '#C62828' for p in pcts]

            bars = ax.barh(names, pcts, color=colors, height=0.5)
            ax.set_title('Today\'s Performance (%)', fontsize=11,
                        fontweight='bold', color='#1B2A4A')
            ax.axvline(x=0, color='#D0D5DD', linewidth=0.5)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            buffer = io.BytesIO()
            fig.savefig(buffer, format='png', dpi=200,
                       bbox_inches='tight', facecolor='white')
            plt.close(fig)
            buffer.seek(0)

            charts.append({
                'title': '今日全球指數表現',
                'base64': base64.b64encode(buffer.read()).decode('utf-8'),
            })

        analysis['charts'] = charts
        logger.info(f"  已生成 {len(charts)} 張圖表")
        return analysis

    def step4_generate_pdf(self, analysis):
        """步驟四：生成 PDF"""
        logger.info("步驟 4/5：生成 PDF 報告...")

        date_str = datetime.date.today().strftime('%Y%m%d')
        output_path = os.path.join(
            self.output_dir,
            f"macro_daily_{date_str}.pdf"
        )

        # 使用前面定義的 generate_macro_daily 函數
        # generate_macro_daily(analysis, output_path)

        logger.info(f"  PDF 已儲存至：{output_path}")
        return output_path

    def step5_notify(self, pdf_path):
        """步驟五：發送通知/郵件（可選）"""
        email_config = self.config.get('email')
        if not email_config:
            logger.info("步驟 5/5：跳過郵件通知（未配置）")
            return

        logger.info("步驟 5/5：發送郵件通知...")

        try:
            msg = MIMEMultipart()
            msg['From'] = email_config['sender']
            msg['To'] = ', '.join(email_config['recipients'])
            msg['Subject'] = f"宏觀日報 - {datetime.date.today()}"

            body = "附件為今日宏觀經濟日報，請查閱。"
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # 附加 PDF
            with open(pdf_path, 'rb') as f:
                attachment = MIMEBase('application', 'pdf')
                attachment.set_payload(f.read())
                encoders.encode_base64(attachment)
                attachment.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=os.path.basename(pdf_path),
                )
                msg.attach(attachment)

            # 發送
            server = smtplib.SMTP(
                email_config['smtp_server'],
                email_config['smtp_port'],
            )
            server.starttls()
            server.login(email_config['sender'], email_config['password'])
            server.send_message(msg)
            server.quit()

            logger.info("  郵件已發送")

        except Exception as e:
            logger.error(f"  郵件發送失敗：{e}")

    def run(self):
        """執行完整管線"""
        start = datetime.datetime.now()
        logger.info(f"===== 開始執行報告管線 =====")
        logger.info(f"時間：{start.strftime('%Y-%m-%d %H:%M:%S')}")

        try:
            raw_data = self.step1_fetch_data()
            analysis = self.step2_analyze(raw_data)
            analysis = self.step3_create_charts(analysis)
            pdf_path = self.step4_generate_pdf(analysis)
            self.step5_notify(pdf_path)

            elapsed = (datetime.datetime.now() - start).total_seconds()
            logger.info(f"===== 管線完成，耗時 {elapsed:.1f} 秒 =====")
            return pdf_path

        except Exception as e:
            logger.error(f"管線執行失敗：{e}")
            raise

    def _generate_summary(self, data):
        """自動生成每日摘要"""
        if not data.get('indices'):
            return "暫無市場數據。"

        # 簡單規則式摘要
        up_count = sum(1 for idx in data['indices'] if idx['pct'] > 0)
        down_count = sum(1 for idx in data['indices'] if idx['pct'] < 0)
        total = len(data['indices'])

        if up_count > total * 0.7:
            sentiment = "全球市場普遍上漲"
        elif down_count > total * 0.7:
            sentiment = "全球市場普遍下跌"
        else:
            sentiment = "全球市場漲跌互見"

        # 找出最大漲跌幅
        sorted_indices = sorted(data['indices'], key=lambda x: x['pct'])
        best = sorted_indices[-1]
        worst = sorted_indices[0]

        summary = (
            f"{sentiment}。"
            f"表現最佳：{best['name']}（{best['pct']:+.2f}%），"
            f"表現最弱：{worst['name']}（{worst['pct']:+.2f}%）。"
        )

        return summary
```

### 6.2 定時排程（Cron / Schedule）

#### 方法一：系統 Cron（macOS/Linux，最穩定）

```bash
# 編輯 crontab
crontab -e

# 每個交易日早上 9:00 執行（週一至週五）
0 9 * * 1-5 /usr/bin/python3 /path/to/generate_daily_report.py >> /path/to/logs/cron.log 2>&1

# 每週五 18:00 執行週報
0 18 * * 5 /usr/bin/python3 /path/to/generate_weekly_report.py >> /path/to/logs/cron.log 2>&1

# 每月第一個交易日執行月報
0 9 1-7 * 1-5 /usr/bin/python3 /path/to/generate_monthly_report.py >> /path/to/logs/cron.log 2>&1
```

#### 方法二：Python schedule 庫

```python
"""
使用 schedule 庫的 Python 內建排程
適合不需要系統級 cron 的場景
"""

import schedule
import time
import logging

logger = logging.getLogger(__name__)

def job_daily_report():
    """每日報告任務"""
    try:
        pipeline = ReportPipeline(config={
            'output_dir': './reports/daily',
            'report_type': 'daily',
        })
        pipeline.run()
    except Exception as e:
        logger.error(f"每日報告生成失敗：{e}")

def job_weekly_report():
    """每週報告任務"""
    try:
        pipeline = ReportPipeline(config={
            'output_dir': './reports/weekly',
            'report_type': 'weekly',
        })
        pipeline.run()
    except Exception as e:
        logger.error(f"每週報告生成失敗：{e}")

# 排程設定
schedule.every().monday.at("09:00").do(job_daily_report)
schedule.every().tuesday.at("09:00").do(job_daily_report)
schedule.every().wednesday.at("09:00").do(job_daily_report)
schedule.every().thursday.at("09:00").do(job_daily_report)
schedule.every().friday.at("09:00").do(job_daily_report)
schedule.every().friday.at("18:00").do(job_weekly_report)

# 執行排程迴圈
logger.info("排程已啟動，等待任務...")
while True:
    schedule.run_pending()
    time.sleep(60)  # 每分鐘檢查一次
```

#### 方法三：APScheduler（更強大的排程器）

```python
"""
使用 APScheduler 的進階排程
支持 Cron 表達式、持久化、錯過補執行等
"""

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

scheduler = BlockingScheduler()

# 使用 cron 表達式
scheduler.add_job(
    job_daily_report,
    CronTrigger(
        day_of_week='mon-fri',
        hour=9,
        minute=0,
        timezone='Asia/Taipei',
    ),
    id='daily_report',
    name='每日宏觀日報',
    misfire_grace_time=3600,  # 錯過1小時內仍執行
)

scheduler.add_job(
    job_weekly_report,
    CronTrigger(
        day_of_week='fri',
        hour=18,
        minute=0,
        timezone='Asia/Taipei',
    ),
    id='weekly_report',
    name='每週策略週報',
)

scheduler.start()
```

#### 方法四：GitHub Actions（雲端排程）

```yaml
# .github/workflows/daily_report.yml
name: 每日宏觀報告

on:
  schedule:
    # UTC 時間 01:00 = 台北時間 09:00
    - cron: '0 1 * * 1-5'
  workflow_dispatch:  # 手動觸發

jobs:
  generate-report:
    runs-on: ubuntu-latest

    steps:
      - uses: actions/checkout@v4

      - name: 設定 Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: 安裝依賴
        run: |
          pip install reportlab matplotlib pandas yfinance weasyprint

      - name: 生成報告
        run: python scripts/generate_daily_report.py

      - name: 上傳報告
        uses: actions/upload-artifact@v4
        with:
          name: daily-report-${{ github.run_number }}
          path: reports/*.pdf
          retention-days: 30
```

### 6.3 品質控制

```python
"""
PDF 品質控制工具
確保生成的報告符合品質標準
"""

import os
from reportlab.lib.utils import ImageReader


class PDFQualityChecker:
    """PDF 品質檢查器"""

    def __init__(self, max_file_size_mb=10, min_dpi=150):
        self.max_file_size_mb = max_file_size_mb
        self.min_dpi = min_dpi
        self.issues = []

    def check_file_size(self, pdf_path):
        """檢查檔案大小"""
        size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        if size_mb > self.max_file_size_mb:
            self.issues.append(
                f"檔案過大：{size_mb:.1f}MB（上限 {self.max_file_size_mb}MB）"
            )
            return False
        return True

    def check_page_count(self, pdf_path, expected_min=1, expected_max=50):
        """檢查頁數"""
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(pdf_path)
            pages = len(reader.pages)
            if pages < expected_min or pages > expected_max:
                self.issues.append(
                    f"頁數異常：{pages} 頁（預期 {expected_min}-{expected_max}）"
                )
                return False
            return True
        except ImportError:
            # PyPDF2 未安裝，跳過
            return True

    def optimize_images_for_pdf(self, image_paths, target_dpi=150):
        """優化圖片以減小 PDF 檔案大小"""
        try:
            from PIL import Image as PILImage
            optimized = []
            for path in image_paths:
                img = PILImage.open(path)
                # 計算目標解析度
                current_dpi = img.info.get('dpi', (300, 300))
                if current_dpi[0] > target_dpi:
                    ratio = target_dpi / current_dpi[0]
                    new_size = (int(img.width * ratio), int(img.height * ratio))
                    img = img.resize(new_size, PILImage.LANCZOS)
                    img.save(path, optimize=True, quality=85)
                    optimized.append(path)
            return optimized
        except ImportError:
            return []

    def run_all_checks(self, pdf_path):
        """執行所有品質檢查"""
        self.issues = []
        print(f"正在檢查：{pdf_path}")

        self.check_file_size(pdf_path)
        self.check_page_count(pdf_path)

        if self.issues:
            print(f"  發現 {len(self.issues)} 個問題：")
            for issue in self.issues:
                print(f"    - {issue}")
            return False
        else:
            size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
            print(f"  品質檢查通過（{size_mb:.1f}MB）")
            return True


# ===== PDF 檔案大小優化技巧 =====

def optimize_pdf_tips():
    """PDF 檔案大小優化技巧清單"""
    tips = """
    ==========================================
    PDF 檔案大小優化技巧
    ==========================================

    1. 圖表 DPI 控制
       - 螢幕閱讀：150 DPI 足夠
       - 印刷品質：300 DPI
       - 不要超過 300 DPI（增加大小但無明顯品質提升）

    2. 圖片格式選擇
       - 折線圖/柱狀圖：PNG（較小）
       - 照片/複雜圖：JPEG（quality=85）
       - 避免 BMP、TIFF

    3. 字體嵌入
       - 使用 Adobe CID 字體（不嵌入，檔案最小）
       - TrueType 字體自動子集嵌入
       - 減少使用的字體種類（每種字體都會增加檔案大小）

    4. 重複元素優化
       - Logo 等重複圖片只嵌入一次
       - 使用 ReportLab 的圖片快取機制

    5. 壓縮設定
       - ReportLab 預設啟用 zlib 壓縮
       - 可透過 canvas.setCompression(1) 確認壓縮開啟

    6. 頁面數控制
       - 合理使用分頁，避免大量空白頁
       - 緊湊排版減少總頁數
    """
    return tips
```

---

## 附錄 A：常見問題與解決方案

### Q1：中文顯示為方塊或亂碼

```python
# 解決方案：確保正確註冊中文字體
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 方案 A：系統字體（macOS）
pdfmetrics.registerFont(TTFont('ChFont', '/System/Library/Fonts/STHeiti Medium.ttc'))

# 方案 B：Adobe CID 字體
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))  # 繁體

# 然後在 ParagraphStyle 中指定 fontName='ChFont' 或 'MSung-Light'
```

### Q2：表格內容溢出頁面

```python
# 解決方案：限制欄寬 + 文字自動換行
from reportlab.platypus import Paragraph

# 將長文字放入 Paragraph 物件中
data = [
    [Paragraph("很長的文字會自動換行...", style), "100", "200"],
]

# 或設定固定欄寬
table = Table(data, colWidths=[200, 60, 60])
```

### Q3：圖表模糊

```python
# 解決方案：提高 DPI + 使用正確尺寸
fig.savefig(buffer, format='png',
            dpi=300,              # 最少 150，印刷用 300
            bbox_inches='tight',
            pad_inches=0.1)

# 在 PDF 中指定精確尺寸
from reportlab.lib.units import cm
img = Image(buffer, width=15*cm, height=9*cm)
```

### Q4：WeasyPrint 安裝失敗

```bash
# macOS
brew install pango gdk-pixbuf libffi cairo
pip install weasyprint

# 如果仍有問題
brew reinstall pango
export PKG_CONFIG_PATH="/opt/homebrew/lib/pkgconfig"
pip install --no-cache-dir weasyprint
```

### Q5：PDF 檔案太大

```python
# 1. 降低圖表 DPI
fig.savefig(buffer, dpi=150)  # 從 300 降到 150

# 2. 使用 CID 字體（不嵌入）
pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))

# 3. 壓縮圖片
from PIL import Image as PILImage
img = PILImage.open('chart.png')
img.save('chart_optimized.png', optimize=True, quality=85)
```

---

## 附錄 B：工具選型決策樹

```
你的需求是什麼？
│
├── 複雜金融報告（表格、圖表、精密排版）
│   └── → ReportLab + matplotlib
│
├── 設計驅動（經常改版面、有設計師參與）
│   └── → WeasyPrint + Jinja2 + CSS
│
├── 簡單快速（少量數據、固定格式）
│   └── → FPDF2
│
├── 網頁轉 PDF（已有 HTML 頁面）
│   ├── 高保真 → Playwright / Puppeteer
│   └── 輕量級 → WeasyPrint
│
└── 批量生成（大量相似報告）
    └── → ReportLab（效能最佳） 或 FPDF2
```

---

## 附錄 C：推薦學習資源

### 官方文檔
- ReportLab 使用者指南：https://docs.reportlab.com/
- WeasyPrint 文檔：https://doc.courtbouillon.org/weasyprint/stable/
- FPDF2 文檔：https://py-pdf.github.io/fpdf2/
- Jinja2 文檔：https://jinja.palletsprojects.com/

### 教程與範例
- ReportLab + Pandas 自動化報告：https://woteq.com/how-to-create-pdf-reports-automatically-with-python-reportlab-pandas
- WeasyPrint + Jinja2 範例：https://joshkaramuth.com/blog/generate-good-looking-pdfs-weasyprint-jinja2/
- matplotlib 圖表嵌入 ReportLab：https://woteq.com/how-to-generate-charts-with-reportlab-and-matplotlib
- 像素級 PDF 報告：https://medium.com/@alexanderamlani24/creating-pixel-perfect-pdf-reports-using-using-plotly-html-css-weasyprint-and-jinja2-9dafb315f6f8
- Python PDF 工具比較（2025）：https://templated.io/blog/generate-pdfs-in-python-with-libraries/

### 設計參考
- 金融模型配色規範：https://www.wallstreetoasis.com/resources/financial-modeling/financial-model-color-formatting
- 金融儀表板配色：https://www.phoenixstrategy.group/blog/best-color-palettes-for-financial-dashboards
- 金融服務字體設計：https://www.gate39media.com/blog/design-spotlight-fonts-in-financial-services

---

> **本指南總結**
>
> 金融投資 PDF 報告的製作，推薦以 **ReportLab + matplotlib + pandas** 為核心技術棧。
> ReportLab 的 Platypus 引擎提供最強大的排版能力，搭配 matplotlib 的高品質圖表，
> 能夠生成媲美頂級投行研究報告的 PDF 文件。若團隊偏好 HTML/CSS 設計流程，
> **WeasyPrint + Jinja2** 是優秀的替代方案。
>
> 關鍵成功因素：
> 1. 統一的配色與字體規範（建立 Style Guide）
> 2. 模組化的報告組件（封面、表格、圖表可複用）
> 3. 自動化的數據管線（從數據源到 PDF 全自動）
> 4. 嚴格的品質控制（檔案大小、圖表解析度、字體嵌入）
