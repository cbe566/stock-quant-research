# 專業級金融投資PPT製作完整指南

> 整合頂級投行設計風格 + Python自動化製作技術
> 更新日期：2026-03-28

---

## 目錄

- [一、頂級投行PPT設計規範](#一頂級投行ppt設計規範)
  - [1.1 配色方案](#11-配色方案)
  - [1.2 字體選擇](#12-字體選擇)
  - [1.3 版面設計原則](#13-版面設計原則)
  - [1.4 圖表設計規範](#14-圖表設計規範)
- [二、python-pptx 高級技巧](#二python-pptx-高級技巧)
  - [2.1 版面控制](#21-版面控制)
  - [2.2 高級圖形](#22-高級圖形)
  - [2.3 圖表嵌入](#23-圖表嵌入)
  - [2.4 表格美化](#24-表格美化)
  - [2.5 動畫與過渡](#25-動畫與過渡)
- [三、專業金融報告PPT結構](#三專業金融報告ppt結構)
- [四、完整代碼範例](#四完整代碼範例)

---

## 一、頂級投行PPT設計規範

### 1.1 配色方案

#### 1.1.1 各大投行品牌配色

**高盛（Goldman Sachs）**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 主色（深藍） | Ebony Clay | `#22263F` | (34, 38, 63) |
| 品牌藍 | Mariner | `#2178C4` | (33, 120, 196) |
| 輔助藍 | Ruddy Blue | `#64A8F0` | (100, 168, 240) |
| 淺藍 | Light Blue | `#ACD4F1` | (172, 212, 241) |
| 深灰 | Dark Gray | `#231F20` | (35, 31, 32) |
| 中灰 | Medium Gray | `#58575A` | (88, 87, 90) |
| 輔色藍 | Steel Blue | `#7399C6` | (115, 153, 198) |
| 背景白 | White | `#FFFFFF` | (255, 255, 255) |

**摩根士丹利（Morgan Stanley）**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 品牌藍 | MS Blue / Denim | `#187ABA` | (24, 122, 186) |
| 深藍 | Dark Blue | `#002B59` | (0, 43, 89) |
| 灰藍 | Lapis Lazuli | `#336699` | (51, 102, 153) |
| 淺灰 | Light Slate Gray | `#778899` | (119, 136, 153) |
| 黑色 | Black | `#000000` | (0, 0, 0) |
| 白色 | White | `#FFFFFF` | (255, 255, 255) |

**J.P. Morgan**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 主色（深棕） | Dark Brown | `#54301A` | (84, 48, 26) |
| 品牌紅 | JP Red | `#AA1F2E` | (170, 31, 46) |
| 深藍 | Navy Blue | `#1A3E6D` | (26, 62, 109) |
| 灰色 | Charcoal | `#333333` | (51, 51, 51) |
| 白色 | White | `#FFFFFF` | (255, 255, 255) |

**麥肯錫（McKinsey & Company）**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 主色（深藍） | McKinsey Blue | `#24477F` | (36, 71, 127) |
| 輔助藍 | Electric Blue | `#051C2C` | (5, 28, 44) |
| 淺藍 | Light Blue | `#4A90D9` | (74, 144, 217) |
| 白色 | White | `#FFFFFF` | (255, 255, 255) |
| 黑色 | Black | `#000000` | (0, 0, 0) |

**黑石（Blackstone）**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 主色 | Pure Black | `#000000` | (0, 0, 0) |
| 背景 | White | `#FFFFFF` | (255, 255, 255) |
| 輔助灰 | Dark Gray | `#333333` | (51, 51, 51) |
| 輔助灰 | Medium Gray | `#666666` | (102, 102, 102) |

> 黑石以極簡黑白為品牌核心，搭配精緻襯線字體（Kris Sowersby設計的定制字體）

**橋水（Bridgewater Associates）**

| 用途 | 色名 | HEX | RGB |
|------|------|-----|-----|
| 主色（深藍） | Navy | `#003366` | (0, 51, 102) |
| 品牌灰 | Steel Gray | `#555555` | (85, 85, 85) |
| 輔助色 | Teal | `#007A7A` | (0, 122, 122) |
| 背景 | White | `#FFFFFF` | (255, 255, 255) |

#### 1.1.2 金融行業通用專業配色方案

```python
# 金融行業通用配色體系
FINANCE_COLORS = {
    # === 主色系（深藍系 — 信任、專業、穩重）===
    'primary_navy':     '#003366',  # 深海軍藍 — 最常見的金融主色
    'primary_dark':     '#1A2332',  # 近乎黑的深藍 — 標題、重要元素
    'primary_blue':     '#2B5797',  # 中等藍 — 按鈕、強調元素
    'primary_light':    '#4A90D9',  # 亮藍 — 連結、互動元素

    # === 輔助色系 ===
    'secondary_gold':   '#C9A84C',  # 金色 — 高級感、重要標記
    'secondary_teal':   '#007A7A',  # 藍綠 — 差異化、輔助圖表
    'secondary_slate':  '#5B6770',  # 板岩灰 — 次要文字

    # === 中性色系 ===
    'neutral_dark':     '#1E1E1E',  # 深炭灰 — 正文
    'neutral_medium':   '#666666',  # 中灰 — 次要文字
    'neutral_light':    '#999999',  # 淺灰 — 說明文字
    'neutral_border':   '#D9D9D9',  # 邊框灰
    'neutral_bg':       '#F5F5F5',  # 背景灰
    'neutral_white':    '#FFFFFF',  # 白色

    # === 語義色系（數據可視化）===
    'positive':         '#2E7D32',  # 正面/上漲（深綠）
    'positive_light':   '#81C784',  # 正面淺色
    'negative':         '#C62828',  # 負面/下跌（深紅）
    'negative_light':   '#EF9A9A',  # 負面淺色
    'warning':          '#F57F17',  # 警告/橙色
    'neutral_signal':   '#757575',  # 中性/持平

    # === 數據可視化漸變序列（同一指標不同值）===
    'gradient_1':       '#003366',
    'gradient_2':       '#2B5797',
    'gradient_3':       '#4A90D9',
    'gradient_4':       '#7AB8F5',
    'gradient_5':       '#B3D4FC',

    # === 多系列圖表配色（不同指標）===
    'series_1':         '#2B5797',  # 藍
    'series_2':         '#C9A84C',  # 金
    'series_3':         '#007A7A',  # 藍綠
    'series_4':         '#8B4513',  # 棕
    'series_5':         '#6B3FA0',  # 紫
    'series_6':         '#CC6600',  # 橙
}
```

#### 1.1.3 漲跌配色國際慣例

| 市場 | 上漲顏色 | 下跌顏色 | 說明 |
|------|---------|---------|------|
| 美國/歐洲 | 綠色 `#2E7D32` | 紅色 `#C62828` | 國際標準 |
| 中國/日本/韓國 | 紅色 `#C62828` | 綠色 `#2E7D32` | 亞洲慣例 |
| 通用替代方案 | 藍色 `#2B5797` | 紅色 `#C62828` | 適合所有市場 |

> **最佳實踐**：面向國際投資者時，建議使用「藍色=正面、紅色=負面」，避免文化差異造成的誤讀。

---

### 1.2 字體選擇

#### 1.2.1 頂級投行常用字體

| 投行 | 主要字體 | 備註 |
|------|---------|------|
| 高盛 Goldman Sachs | Times New Roman | 金融業專業感的金標準 |
| 摩根士丹利 Morgan Stanley | Times New Roman | 傳統襯線字體 |
| J.P. Morgan | Arial | 簡潔清晰的無襯線字體 |
| 花旗 Citi | Calibri | 現代圓潤風格 |
| 德意志銀行 Deutsche Bank | Calibri | 當代感 |
| 瑞銀 UBS | Helvetica | 瑞士設計精神 |
| 巴克萊 Barclays | Helvetica | 中性優雅 |
| Lazard | Garamond | 古典優雅風格 |
| 麥肯錫 McKinsey | Arial（正文）+ Georgia（標題）| 雙字體搭配 |

#### 1.2.2 中英文字體最佳搭配

| 場景 | 英文字體 | 中文字體 | 備註 |
|------|---------|---------|------|
| 方案A（現代感） | Calibri | 微軟雅黑 | 最安全的選擇，Windows全兼容 |
| 方案B（專業感） | Arial | 思源黑體 | 開源免費，跨平台一致 |
| 方案C（蘋果生態） | Helvetica Neue | 蘋方-繁 | macOS最佳呈現 |
| 方案D（古典感） | Garamond / Georgia | 思源宋體 | 適合正式研究報告 |
| 方案E（數據導向） | Consolas / Menlo | 等寬字體 | 適合展示代碼或數字表格 |

#### 1.2.3 字號規範體系

```
投行PPT字號標準：
┌─────────────────────────────────────────────┐
│ 元素          │ 字號(pt)  │ 字重      │ 備註              │
├─────────────────────────────────────────────┤
│ 封面主標題     │ 28-36    │ Bold      │ 簡短有力           │
│ 封面副標題     │ 18-22    │ Regular   │ 補充說明           │
│ 頁面標題       │ 20-24    │ Bold      │ 每頁一個核心觀點    │
│ 行動標題       │ 14-16    │ Bold      │ McKinsey風格結論句  │
│ 正文          │ 11-14    │ Regular   │ 投行標準正文尺寸    │
│ 圖表標題       │ 12-14    │ Bold      │ 簡述圖表內容       │
│ 圖表標籤/註釋  │ 8-10     │ Regular   │ 數據標籤、軸標籤    │
│ 腳註/來源      │ 7-8      │ Light     │ 頁面底部           │
│ 頁碼          │ 8-9      │ Regular   │ 右下角             │
│ 表格內容       │ 9-11     │ Regular   │ 根據數據密度調整    │
│ 表格表頭       │ 10-12    │ Bold      │ 深色背景白字       │
│ KPI大數字      │ 36-48    │ Bold      │ 儀表板關鍵指標     │
│ KPI說明文字    │ 10-12    │ Regular   │ 大數字下方說明     │
└─────────────────────────────────────────────┘
```

---

### 1.3 版面設計原則

#### 1.3.1 McKinsey/投行PPT的標準版面結構

每一頁投影片都遵循嚴格的三段式結構：

```
┌─────────────────────────────────────────────────────────┐
│ [Logo]          行動標題（Action Title）        [頁碼]  │  ← 頂部區：結論句
├─────────────────────────────────────────────────────────┤
│                                                         │
│                                                         │
│                  主體內容區域                             │  ← 主體區：圖表/表格/文字
│              （圖表、表格、數據）                         │
│                                                         │
│                                                         │
├─────────────────────────────────────────────────────────┤
│ 來源：Bloomberg, 公司年報  │ 機密聲明    │ 日期         │  ← 底部區：來源與聲明
└─────────────────────────────────────────────────────────┘
```

**核心原則：**

1. **一頁一個核心觀點（One Slide, One Message）**
   - 行動標題用完整句子表達結論，而非描述性標題
   - 錯誤示範：「營收分析」
   - 正確示範：「公司營收連續三季度加速增長，預計FY2026達到150億美元」

2. **金字塔原則（Pyramid Principle）**
   - 結論先行，細節在後
   - 從最重要的訊息開始，逐層展開支撐論據
   - 垂直邏輯流（Vertical Flow）：從上到下，從核心到細節

3. **MECE原則**
   - 彼此獨立（Mutually Exclusive）：每頁講不同的事
   - 完全窮盡（Collectively Exhaustive）：合在一起涵蓋全部

#### 1.3.2 留白與資訊密度

```
投行PPT留白規範：
┌──────────────────────────────────────┐
│ 邊距    │ 建議值        │ 備註      │
├──────────────────────────────────────┤
│ 上邊距   │ 0.5-0.7 英寸  │ Logo區域  │
│ 下邊距   │ 0.3-0.5 英寸  │ 來源區域  │
│ 左邊距   │ 0.5-0.7 英寸  │           │
│ 右邊距   │ 0.5-0.7 英寸  │           │
│ 元素間距  │ 0.1-0.2 英寸  │ 最小間距  │
└──────────────────────────────────────┘
```

- **資訊密度**：投行PPT通常比科技業PPT密度更高（每頁內容更多），但仍需留白
- **70%原則**：內容面積不超過投影片面積的70%
- **對齊**：所有元素必須嚴格網格對齊，絕不「大概差不多」

#### 1.3.3 頁首頁尾規範

| 位置 | 內容 | 對齊方式 |
|------|------|---------|
| 左上角 | 公司Logo（高度約0.4英寸） | 左對齊 |
| 頂部中間 | 行動標題或章節標題 | 居中或左對齊 |
| 右上角 | 頁碼或日期 | 右對齊 |
| 左下角 | 數據來源（Source: ...） | 左對齊 |
| 右下角 | 機密聲明 / 頁碼 | 右對齊 |
| 底部分隔 | 細線（0.5pt, 灰色）分隔正文與腳註 | 全寬 |

---

### 1.4 圖表設計規範

#### 1.4.1 各類圖表最佳實踐

**折線圖（股價走勢、指數比較）**
- 線寬：1.5-2.5pt
- 最多不超過5條線，否則改用小多圖（Small Multiples）
- 主要線條用品牌深藍，次要線條用灰色系
- 關鍵轉折點加標註（callout）
- Y軸不一定從0開始（金融圖表的慣例）
- 加入陰影區域標示事件（如衰退期）

**柱狀圖/條形圖（營收、盈利比較）**
- 柱寬與間距比約 2:1
- 同系列用同色不同深淺，不同系列用不同色
- 數據標籤直接標在柱子上方，減少視覺追蹤
- 基線（0線）加粗強調
- 負值使用紅色/淺色

**圓餅圖/甜甜圈圖（配置比例、市場份額）**
- 最多6-7個切片，其餘歸入「其他」
- 從12點鐘方向開始，最大塊順時針排列
- 甜甜圈圖中心放總值或標題
- 標籤放在圖外用引導線
- 強調項微微突出（explode）

**瀑布圖（財務分析、營收拆解）**
- 增加項用藍色/綠色，減少項用紅色
- 起點和終點用深色強調
- 每個柱子標註具體數值
- 使用連接線表示累計效果

**散點圖（風險/回報分析）**
- X軸：風險指標（波動率/Beta）
- Y軸：回報指標
- 加入象限分隔線和標籤
- 氣泡大小可表示第三維度（如市值）
- 標註關鍵資產名稱

**熱力圖（相關性矩陣、因子暴露）**
- 使用發散色階（Diverging Colormap）：紅-白-藍 或 紅-黃-綠
- 正相關/正暴露偏藍，負相關/負暴露偏紅
- 對角線可用灰色或留空
- 顯示數值在格子中

**K線圖（技術分析）**
- 陽線（漲）：空心或藍/綠
- 陰線（跌）：實心或紅
- 成交量柱在下方，與K線同色
- 均線用不同顏色虛線
- 支撐/壓力位用水平虛線標註

#### 1.4.2 圖表通用設計規則

```
投行圖表「必做」清單：
✓ 每張圖表必須有標題（描述結論而非描述圖表）
✓ 標註數據來源和日期
✓ 清除不必要的網格線（保留淡灰色水平參考線即可）
✓ 數據標籤直接標在數據點旁（減少圖例依賴）
✓ 使用一致的數字格式（千分位、百分比、貨幣符號）
✓ 圖表配色與整體PPT風格統一
✓ 關鍵數據點加粗或放大
✓ 加入「So What」標註 — 指出圖表的核心洞察

投行圖表「禁忌」清單：
✗ 3D效果（永遠不要用）
✗ 過多動畫
✗ 過於鮮豔的配色
✗ 裝飾性元素（clipart、圖片背景）
✗ 未標註來源的數據
✗ Y軸刻度不均勻
✗ 圖例過多（超過5個）
```

---

## 二、python-pptx 高級技巧

### 2.1 版面控制

#### 2.1.1 EMU單位系統

python-pptx 使用 EMU（English Metric Units）作為基礎單位。常用轉換工具：

```python
from pptx.util import Inches, Cm, Pt, Emu

# 單位轉換
# 1英寸 = 914400 EMU
# 1釐米 = 360000 EMU
# 1點(pt) = 12700 EMU

width = Inches(10)    # 10英寸
height = Cm(19.05)    # 19.05釐米
font_size = Pt(14)    # 14點

# 標準16:9投影片尺寸
SLIDE_WIDTH = Inches(13.333)   # 33.867 cm
SLIDE_HEIGHT = Inches(7.5)     # 19.05 cm

# 標準4:3投影片尺寸
SLIDE_WIDTH_43 = Inches(10)
SLIDE_HEIGHT_43 = Inches(7.5)
```

#### 2.1.2 自定義母版（Slide Master）

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def create_presentation_with_master():
    """建立帶有自定義母版設定的簡報"""
    prs = Presentation()

    # 設定投影片大小為16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 存取投影片母版
    slide_master = prs.slide_masters[0]

    # 存取版面配置
    # 常見的版面配置索引：
    # 0 = Title Slide（標題投影片）
    # 1 = Title and Content（標題與內容）
    # 5 = Blank（空白）
    # 6 = Content with Caption（帶標題的內容）

    for i, layout in enumerate(prs.slide_layouts):
        print(f"Layout {i}: {layout.name}")

    return prs
```

#### 2.1.3 佔位符（Placeholder）管理

```python
def explore_placeholders(slide):
    """探索投影片中所有佔位符的位置和大小"""
    for shape in slide.placeholders:
        print(f"idx: {shape.placeholder_format.idx}")
        print(f"  name: {shape.name}")
        print(f"  type: {shape.placeholder_format.type}")
        print(f"  position: left={shape.left}, top={shape.top}")
        print(f"  size: width={shape.width}, height={shape.height}")
        print()

def set_placeholder_text(slide, idx, text, font_name='Calibri',
                         font_size=14, bold=False, color='#1E1E1E'):
    """設定佔位符文字及格式"""
    ph = slide.placeholders[idx]
    ph.text = text
    for paragraph in ph.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color[1:])
```

#### 2.1.4 精確位置控制工具函式

```python
class SlideGrid:
    """投影片網格系統 — 協助精確佈局"""

    def __init__(self, slide_width=Inches(13.333), slide_height=Inches(7.5),
                 margin_left=Inches(0.6), margin_right=Inches(0.6),
                 margin_top=Inches(0.5), margin_bottom=Inches(0.4)):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.margin_left = margin_left
        self.margin_right = margin_right
        self.margin_top = margin_top
        self.margin_bottom = margin_bottom

        # 可用區域
        self.content_width = slide_width - margin_left - margin_right
        self.content_height = slide_height - margin_top - margin_bottom

    def get_position(self, col, row, total_cols, total_rows,
                     gap_x=Inches(0.15), gap_y=Inches(0.15)):
        """
        計算網格中某個單元格的位置和大小
        col, row: 從0開始的索引
        total_cols, total_rows: 網格的列數和行數
        """
        cell_width = (self.content_width - gap_x * (total_cols - 1)) // total_cols
        cell_height = (self.content_height - gap_y * (total_rows - 1)) // total_rows

        left = self.margin_left + col * (cell_width + gap_x)
        top = self.margin_top + row * (cell_height + gap_y)

        return left, top, cell_width, cell_height
```

---

### 2.2 高級圖形

#### 2.2.1 漸變填充（Gradient Fill）

```python
from pptx.oxml.ns import qn
from lxml import etree

def add_gradient_rectangle(slide, left, top, width, height,
                           color1='#003366', color2='#4A90D9', angle=270):
    """
    新增漸變填充的矩形
    angle: 270=從上到下, 0=從左到右, 90=從下到上
    """
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    shape.line.fill.background()  # 移除邊框

    # 設定漸變填充（使用XML操作）
    spPr = shape._element.spPr

    # 移除現有填充
    for child in list(spPr):
        if child.tag.endswith('solidFill') or child.tag.endswith('gradFill'):
            spPr.remove(child)

    # 建立漸變填充XML
    gradFill = etree.SubElement(spPr, qn('a:gradFill'))
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

    # 起始顏色
    gs1 = etree.SubElement(gsLst, qn('a:gs'))
    gs1.set('pos', '0')
    srgb1 = etree.SubElement(etree.SubElement(gs1, qn('a:srgbClr')), qn('a:srgbClr'))
    gs1_clr = etree.SubElement(gs1, qn('a:srgbClr'))
    gs1_clr.set('val', color1[1:])
    # 移除多餘的子元素
    for child in list(gs1):
        if child.tag != qn('a:srgbClr') or child.get('val') != color1[1:]:
            gs1.remove(child)

    # 結束顏色
    gs2 = etree.SubElement(gsLst, qn('a:gs'))
    gs2.set('pos', '100000')
    gs2_clr = etree.SubElement(gs2, qn('a:srgbClr'))
    gs2_clr.set('val', color2[1:])

    # 線性漸變方向
    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', str(angle * 60000))  # 角度轉EMU角度
    lin.set('scaled', '1')

    return shape
```

#### 2.2.2 圓角矩形與陰影

```python
from pptx.enum.shapes import MSO_SHAPE

def add_rounded_card(slide, left, top, width, height,
                     fill_color='#FFFFFF', border_color='#D9D9D9',
                     border_width=Pt(0.75), shadow=True):
    """
    新增圓角卡片（投行KPI卡片常用風格）
    """
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )

    # 填充顏色
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(fill_color[1:])

    # 邊框
    shape.line.color.rgb = RGBColor.from_string(border_color[1:])
    shape.line.width = border_width

    # 陰影效果（透過XML操作）
    if shadow:
        spPr = shape._element.spPr
        effectLst = etree.SubElement(spPr, qn('a:effectLst'))
        outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
        outerShdw.set('blurRad', '50800')    # 模糊半徑
        outerShdw.set('dist', '25400')        # 偏移距離
        outerShdw.set('dir', '5400000')       # 方向（下方）
        outerShdw.set('rotWithShape', '0')

        srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
        srgbClr.set('val', '000000')
        alpha = etree.SubElement(srgbClr, qn('a:alpha'))
        alpha.set('val', '25000')  # 25%透明度

    return shape
```

#### 2.2.3 透明度控制

```python
def set_shape_transparency(shape, transparency_pct):
    """
    設定形狀填充的透明度
    transparency_pct: 0-100 的整數
    """
    # 確保有填充
    fill = shape.fill
    if fill.type is None:
        fill.solid()

    # 透過XML設定alpha值
    solidFill = shape._element.spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        srgbClr = solidFill.find(qn('a:srgbClr'))
        if srgbClr is not None:
            # 移除舊的alpha
            for alpha in srgbClr.findall(qn('a:alpha')):
                srgbClr.remove(alpha)
            # 新增alpha（100000 = 100%不透明）
            alpha = etree.SubElement(srgbClr, qn('a:alpha'))
            alpha.set('val', str((100 - transparency_pct) * 1000))
```

#### 2.2.4 連接線與箭頭

```python
from pptx.enum.shapes import MSO_CONNECTOR_TYPE

def add_connector(slide, start_x, start_y, end_x, end_y,
                  color='#999999', width=Pt(1)):
    """新增連接線"""
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT,
        start_x, start_y, end_x, end_y
    )
    connector.line.color.rgb = RGBColor.from_string(color[1:])
    connector.line.width = width
    return connector
```

---

### 2.3 圖表嵌入

#### 2.3.1 python-pptx 原生圖表 vs matplotlib 圖片

| 面向 | python-pptx 原生圖表 | matplotlib 圖片嵌入 |
|------|---------------------|-------------------|
| 可編輯性 | 在PPT中可直接編輯 | 嵌入後為圖片，不可編輯 |
| 圖表類型 | 柱狀、折線、圓餅、散點等基本類型 | 幾乎無限制 |
| 美觀度 | 受PPT圖表引擎限制 | 完全自定義 |
| 複雜度 | API相對繁瑣 | matplotlib語法更靈活 |
| 建議用途 | 簡單圖表、需要客戶修改的圖表 | 複雜圖表、精緻視覺化 |

**推薦策略**：簡單圖表用原生，複雜圖表用matplotlib生成透明PNG嵌入。

#### 2.3.2 python-pptx 原生圖表範例

```python
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def add_line_chart(slide, left, top, width, height,
                   categories, series_data, title=None):
    """
    新增折線圖
    categories: ['Q1', 'Q2', 'Q3', 'Q4']
    series_data: {'營收': [100, 120, 115, 140], '利潤': [20, 25, 22, 30]}
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories

    for name, values in series_data.items():
        chart_data.add_series(name, values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    # 設定標題
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True

    # 圖例位置
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(9)

    # 設定Y軸格式
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.tick_labels.font.size = Pt(8)

    # 設定X軸格式
    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = Pt(8)
    category_axis.has_major_gridlines = False

    # 設定線條顏色
    colors = ['#2B5797', '#C9A84C', '#007A7A', '#C62828', '#6B3FA0']
    for i, series in enumerate(chart.series):
        series.format.line.color.rgb = RGBColor.from_string(colors[i % len(colors)][1:])
        series.format.line.width = Pt(2)
        series.smooth = False

    return chart


def add_bar_chart(slide, left, top, width, height,
                  categories, series_data, title=None):
    """新增柱狀圖"""
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_data.items():
        chart_data.add_series(name, values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title

    # 柱狀圖配色
    colors = ['#2B5797', '#C9A84C', '#007A7A']
    for i, series in enumerate(chart.series):
        fill = series.format.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)][1:])

    return chart


def add_pie_chart(slide, left, top, width, height,
                  categories, values, title=None):
    """新增圓餅圖"""
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series('', values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, left, top, width, height, chart_data
    )
    chart = chart_frame.chart

    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title

    # 顯示數據標籤
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.number_format = '0.0%'
    data_labels.font.size = Pt(9)
    data_labels.show_percentage = True
    data_labels.show_category_name = True

    return chart
```

#### 2.3.3 matplotlib 生成高品質透明圖表並嵌入

```python
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # 非互動模式
import numpy as np
from io import BytesIO
from pptx.util import Inches

# 設定全域風格（投行風格）
def setup_finance_style():
    """設定投行風格的matplotlib全域樣式"""
    plt.rcParams.update({
        # 字體
        'font.family': ['Calibri', 'Arial', 'Helvetica Neue', 'sans-serif'],
        'font.size': 10,
        'axes.titlesize': 12,
        'axes.titleweight': 'bold',
        'axes.labelsize': 10,

        # 顏色
        'axes.facecolor': 'none',        # 透明背景
        'figure.facecolor': 'none',      # 透明背景
        'axes.edgecolor': '#D9D9D9',
        'axes.labelcolor': '#333333',
        'xtick.color': '#666666',
        'ytick.color': '#666666',
        'text.color': '#1E1E1E',

        # 網格
        'axes.grid': True,
        'grid.color': '#E8E8E8',
        'grid.linewidth': 0.5,
        'grid.alpha': 0.7,

        # 刻度
        'xtick.major.size': 0,
        'ytick.major.size': 0,

        # 圖框
        'axes.spines.top': False,
        'axes.spines.right': False,

        # DPI
        'figure.dpi': 200,
        'savefig.dpi': 300,
    })

setup_finance_style()


def fig_to_pptx_image(fig, slide, left, top, width, height=None):
    """
    將matplotlib圖表轉為透明PNG並嵌入投影片
    """
    buf = BytesIO()
    fig.savefig(buf, format='png', transparent=True,
                bbox_inches='tight', pad_inches=0.1)
    buf.seek(0)
    plt.close(fig)

    if height:
        slide.shapes.add_picture(buf, left, top, width, height)
    else:
        slide.shapes.add_picture(buf, left, top, width=width)

    return slide


def create_stock_chart(dates, prices, title='股價走勢',
                       ma_periods=[20, 50, 200]):
    """
    生成投行風格的股價走勢圖
    """
    fig, ax = plt.subplots(figsize=(10, 5))

    # 主線
    ax.plot(dates, prices, color='#2B5797', linewidth=1.5, label='收盤價')

    # 均線
    ma_colors = {'20': '#C9A84C', '50': '#007A7A', '200': '#C62828'}
    for period in ma_periods:
        if len(prices) >= period:
            ma = np.convolve(prices, np.ones(period)/period, mode='valid')
            ax.plot(dates[period-1:], ma, linewidth=1,
                   linestyle='--', color=ma_colors.get(str(period), '#999'),
                   label=f'MA{period}', alpha=0.8)

    ax.set_title(title, fontsize=13, fontweight='bold', color='#1A2332', pad=15)
    ax.legend(loc='upper left', framealpha=0.9, fontsize=8)
    ax.set_ylabel('價格', fontsize=9)

    # 格式化Y軸
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))

    fig.tight_layout()
    return fig
```

---

### 2.4 表格美化

#### 2.4.1 自定義專業表格

```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def add_professional_table(slide, left, top, width, height,
                           headers, data, col_widths=None,
                           header_color='#003366',
                           stripe_color='#F5F7FA',
                           highlight_rules=None):
    """
    新增專業投行風格表格

    參數：
        headers: ['指標', 'FY2024', 'FY2025E', 'FY2026E']
        data: [['營收(億)', 100, 120, 145], ['淨利(億)', 15, 20, 28], ...]
        col_widths: [Inches(2.5), Inches(1.8), ...] 各列寬度
        highlight_rules: 條件格式規則函數
    """
    rows = len(data) + 1  # +1 表頭
    cols = len(headers)

    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # 設定列寬
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w

    # === 表頭 ===
    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(header)

        # 表頭背景色
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(header_color[1:])

        # 表頭文字格式
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
            paragraph.font.name = 'Calibri'
            paragraph.alignment = PP_ALIGN.CENTER

        # 垂直置中
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # === 數據行 ===
    for i, row_data in enumerate(data):
        for j, value in enumerate(row_data):
            cell = table.cell(i + 1, j)

            # 格式化數值
            if isinstance(value, float):
                cell.text = f'{value:,.2f}'
            elif isinstance(value, int):
                cell.text = f'{value:,}'
            else:
                cell.text = str(value)

            # 斑馬紋（偶數行上底色）
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(stripe_color[1:])
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

            # 文字格式
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.name = 'Calibri'
                paragraph.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)

                # 第一欄左對齊，其餘右對齊
                if j == 0:
                    paragraph.alignment = PP_ALIGN.LEFT
                else:
                    paragraph.alignment = PP_ALIGN.RIGHT

            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            # === 條件格式 ===
            if highlight_rules and j > 0:
                try:
                    num_val = float(value) if not isinstance(value, str) else None
                    if num_val is not None:
                        color = highlight_rules(num_val, i, j)
                        if color:
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.color.rgb = RGBColor.from_string(color[1:])
                except (ValueError, TypeError):
                    pass

    # 移除預設表格樣式的邊框（使用XML）
    tbl = table._tbl
    tbl.attrib.clear()

    return table


def positive_negative_rule(value, row, col):
    """條件格式規則：正數綠色、負數紅色"""
    if value > 0:
        return '#2E7D32'  # 深綠
    elif value < 0:
        return '#C62828'  # 深紅
    return None


def percentage_change_rule(value, row, col):
    """百分比變化的條件格式"""
    if value > 10:
        return '#1B5E20'   # 深綠（強勢）
    elif value > 0:
        return '#2E7D32'   # 綠色（正面）
    elif value > -10:
        return '#C62828'   # 紅色（負面）
    else:
        return '#B71C1C'   # 深紅（弱勢）
```

#### 2.4.2 合併儲存格

```python
def merge_cells_example(table):
    """合併儲存格示範"""
    # 合併第一行的第2-3列（水平合併）
    cell_a = table.cell(0, 1)
    cell_b = table.cell(0, 2)
    cell_a.merge(cell_b)

    # 合併第一列的第2-4行（垂直合併）
    cell_c = table.cell(1, 0)
    cell_d = table.cell(3, 0)
    cell_c.merge(cell_d)

    # 合併後，格式以左上角（merge origin）為準
    cell_a.text = "合併後的標題"
```

---

### 2.5 動畫與過渡

python-pptx **不支援**原生動畫和投影片過渡效果的API。但有以下替代方案：

```python
# === 方案一：透過XML注入過渡效果 ===
from pptx.oxml.ns import qn
from lxml import etree

def add_slide_transition(slide, transition_type='fade', duration=500):
    """
    為投影片新增過渡效果（XML注入方式）
    transition_type: 'fade', 'push', 'wipe', 'dissolve'
    duration: 毫秒
    """
    transition_map = {
        'fade': 'fade',
        'push': 'push',
        'wipe': 'wipe',
        'dissolve': 'dissolve',
    }

    sld = slide._element
    # 建立 transition 元素
    ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    transition = etree.SubElement(sld, f'{{{ns}}}transition')
    transition.set('spd', 'med')
    transition.set('advClick', '1')

    # 新增具體過渡類型
    trans_elem = etree.SubElement(transition,
                                  f'{{{ns}}}{transition_map.get(transition_type, "fade")}')

    return slide

# === 方案二：用漸進式顯示模擬動畫（建立多頁） ===
def create_progressive_reveal(prs, data_list, title):
    """
    建立漸進式揭示效果（每頁多顯示一個元素）
    模擬「逐步出現」的動畫效果
    """
    for i in range(1, len(data_list) + 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版面

        # 標題
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(0.3),
                                          Inches(12), Inches(0.6))
        tf = txBox.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True

        # 逐步添加內容
        for j in range(i):
            y_pos = Inches(1.2 + j * 0.6)
            txBox = slide.shapes.add_textbox(Inches(1), y_pos,
                                              Inches(11), Inches(0.5))
            tf = txBox.text_frame
            tf.text = data_list[j]
            tf.paragraphs[0].font.size = Pt(14)

            # 最新出現的元素用深色，之前的用灰色
            if j == i - 1:
                tf.paragraphs[0].font.color.rgb = RGBColor(0x1A, 0x23, 0x32)
                tf.paragraphs[0].font.bold = True
            else:
                tf.paragraphs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)
```

---

## 三、專業金融報告PPT結構

### 3.1 股票研究報告PPT結構

```
投影片順序與內容：

第1頁 — 封面
  ├── 公司名稱 + 股票代碼
  ├── 投資評級（買入/持有/賣出）+ 目標價
  ├── 報告日期 + 分析師姓名
  └── 公司Logo + 研究機構Logo

第2頁 — 投資摘要（Executive Summary）
  ├── 3-5個核心投資觀點（子彈點）
  ├── 關鍵財務指標摘要表
  ├── 目標價與當前價差距（上行/下行空間）
  └── 評級理由一句話總結

第3頁 — 公司概覽
  ├── 業務描述與產品線（甜甜圈圖：營收構成）
  ├── 行業地位與市場份額
  ├── 競爭優勢（護城河分析）
  └── 管理團隊簡介

第4頁 — 行業分析
  ├── 行業規模與增長趨勢（折線圖）
  ├── 競爭格局（波特五力或市場份額圖）
  ├── 行業驅動因子
  └── 監管環境

第5頁 — 財務分析（營收與利潤）
  ├── 營收趨勢與增長率（柱狀+折線組合圖）
  ├── 毛利率/營業利潤率/淨利率趨勢
  ├── 同業比較表
  └── 關鍵增長驅動因子

第6頁 — 財務分析（資產負債表與現金流）
  ├── 資產結構變化
  ├── 負債水平與償債能力
  ├── 自由現金流趨勢
  └── ROE/ROIC分解（杜邦分析）

第7頁 — 估值分析
  ├── DCF模型假設與結果
  ├── 相對估值（P/E、P/B、EV/EBITDA vs 同業）
  ├── 歷史估值區間
  └── 目標價推導表格

第8頁 — 技術分析
  ├── K線圖（含均線）
  ├── MACD / RSI 指標
  ├── 支撐壓力位標註
  └── 成交量分析

第9頁 — 籌碼分析（台股適用）
  ├── 三大法人買賣超趨勢
  ├── 千張大戶持股變化
  ├── 融資融券餘額
  └── 券商分點追蹤

第10頁 — 風險因素
  ├── 風險矩陣（影響程度 x 發生機率）
  ├── 3-5個主要風險因素詳述
  └── 敏感度分析

第11頁 — 投資建議與目標價
  ├── 最終評級與目標價
  ├── 建議的進場/出場價位
  ├── 時間框架
  └── 倉位建議

第12頁 — 附錄
  ├── 詳細財務報表（損益表/資產負債表/現金流量表）
  ├── 關鍵假設列表
  └── 術語解釋

第13頁 — 免責聲明
  ├── 研究報告聲明
  ├── 利益衝突揭露
  └── 版權聲明
```

### 3.2 量化策略報告PPT結構

```
第1頁 — 封面
  ├── 策略名稱
  ├── 回測期間
  ├── 核心績效摘要（年化報酬、夏普比率、最大回撤）
  └── 日期

第2頁 — 策略概述
  ├── 投資理念與邏輯
  ├── 策略流程圖（因子篩選 → 組合構建 → 風險管理）
  ├── 適用市場與資產類別
  └── 調倉頻率

第3頁 — 回測績效總覽
  ├── 淨值曲線（vs 基準指數）
  ├── 年度報酬率柱狀圖
  ├── 月度報酬率熱力圖
  └── 關鍵績效指標表（年化報酬、波動率、夏普、Calmar等）

第4頁 — 風險分析
  ├── 回撤曲線
  ├── VaR / CVaR 統計
  ├── 壓力測試結果（2008金融危機、2020疫情等）
  └── 滾動夏普比率

第5頁 — 因子分析
  ├── 因子暴露雷達圖
  ├── 因子IC時間序列
  ├── 因子收益歸因（Brinson模型）
  └── 因子相關性矩陣

第6頁 — 持倉分析
  ├── 行業分布（圓餅圖）
  ├── 個股權重分布（Treemap）
  ├── 持倉集中度指標
  └── 換手率趨勢

第7頁 — 交易統計
  ├── 勝率 / 盈虧比
  ├── 平均持有天數
  ├── 交易頻率
  ├── 滑點與交易成本影響
  └── 最大連續虧損次數

第8頁 — 總結與建議
  ├── 策略優劣勢分析
  ├── 適用場景與限制
  ├── 優化方向
  └── 實盤部署建議
```

### 3.3 宏觀研究報告PPT結構

```
第1頁 — 封面
第2頁 — 全球經濟儀表板（一頁KPI摘要）
  ├── GDP增長率（主要經濟體）
  ├── 通膨率
  ├── 利率水準
  ├── 主要股指表現
  └── VIX指數

第3頁 — 景氣循環判斷
  ├── 當前所處階段（擴張/高峰/收縮/谷底）
  ├── 領先指標趨勢
  ├── PMI走勢
  └── 殖利率曲線

第4頁 — 美國經濟分析
第5頁 — 中國經濟分析
第6頁 — 歐洲/日本經濟分析
第7頁 — 央行政策動向
  ├── Fed利率路徑預測
  ├── 各央行QT進程
  └── 政策分歧分析

第8頁 — 資金流向與流動性
  ├── 全球資金流向圖
  ├── 美元指數走勢
  ├── 信用利差
  └── 跨境資本流動

第9頁 — 各資產類別展望
  ├── 股票（美/中/歐/日）
  ├── 債券（利率/信用）
  ├── 商品（黃金/原油）
  ├── 匯率
  └── 另類投資

第10頁 — 風險地圖
  ├── 地緣政治風險
  ├── 金融市場風險
  ├── 黑天鵝情境分析
  └── 尾部風險對沖建議

第11頁 — 投資建議
  ├── 資產配置建議（保守/均衡/積極）
  ├── 板塊輪動建議
  ├── 具體標的推薦
  └── 時間框架與風控
```

---

## 四、完整代碼範例

### 4.1 專業封面頁模板

```python
"""
範例1：專業投行風格封面頁
可直接運行，生成帶漸變背景、Logo區域、標題的封面
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from datetime import datetime


def create_cover_slide(prs, title, subtitle, author, date_str=None,
                       rating=None, target_price=None, ticker=None):
    """
    建立專業投行風格封面頁

    參數：
        title: 主標題（如公司名稱）
        subtitle: 副標題（如「深度研究報告」）
        author: 分析師姓名
        date_str: 日期字串
        rating: 投資評級（如「買入」）
        target_price: 目標價
        ticker: 股票代碼
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版面

    if date_str is None:
        date_str = datetime.now().strftime('%Y年%m月%d日')

    # === 1. 深藍漸變背景（上半部）===
    bg_top = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(4.5)
    )
    bg_top.line.fill.background()

    # 漸變填充
    spPr = bg_top._element.spPr
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        spPr.remove(solidFill)

    gradFill = etree.SubElement(spPr, qn('a:gradFill'))
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

    gs1 = etree.SubElement(gsLst, qn('a:gs'))
    gs1.set('pos', '0')
    c1 = etree.SubElement(gs1, qn('a:srgbClr'))
    c1.set('val', '1A2332')

    gs2 = etree.SubElement(gsLst, qn('a:gs'))
    gs2.set('pos', '100000')
    c2 = etree.SubElement(gs2, qn('a:srgbClr'))
    c2.set('val', '2B5797')

    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', '5400000')  # 從左到右
    lin.set('scaled', '1')

    # === 2. 金色裝飾線 ===
    line_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.0), Inches(2), Pt(3)
    )
    line_shape.fill.solid()
    line_shape.fill.fore_color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
    line_shape.line.fill.background()

    # === 3. 主標題 ===
    title_box = slide.shapes.add_textbox(
        Inches(0.8), Inches(1.5), Inches(10), Inches(1.2)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]

    # 如果有股票代碼，加在標題前面
    if ticker:
        run_ticker = p.add_run()
        run_ticker.text = f'{ticker}  '
        run_ticker.font.size = Pt(20)
        run_ticker.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
        run_ticker.font.name = 'Calibri'

    run_title = p.add_run()
    run_title.text = title
    run_title.font.size = Pt(36)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run_title.font.name = 'Calibri'

    # === 4. 副標題 ===
    subtitle_box = slide.shapes.add_textbox(
        Inches(0.8), Inches(3.3), Inches(10), Inches(0.6)
    )
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = subtitle
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0xAC, 0xD4, 0xF1)
    run.font.name = 'Calibri'

    # === 5. 評級與目標價（如果有）===
    if rating or target_price:
        rating_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(9.5), Inches(1.5), Inches(3.2), Inches(2)
        )
        rating_box.fill.solid()
        rating_box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # 設定透明度
        srgbClr = rating_box._element.spPr.find(f'.//{qn("a:srgbClr")}')
        if srgbClr is not None:
            alpha = etree.SubElement(srgbClr, qn('a:alpha'))
            alpha.set('val', '20000')  # 20%不透明

        rating_box.line.fill.background()

        tf = rating_box.text_frame
        tf.word_wrap = True

        if rating:
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = f'投資評級：{rating}'
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        if target_price:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = f'目標價：{target_price}'
            run.font.size = Pt(22)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)

    # === 6. 底部資訊區（淺灰背景）===
    bg_bottom = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, Inches(5.5), prs.slide_width, Inches(2)
    )
    bg_bottom.fill.solid()
    bg_bottom.fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
    bg_bottom.line.fill.background()

    # 分析師與日期
    info_box = slide.shapes.add_textbox(
        Inches(0.8), Inches(5.8), Inches(8), Inches(1.2)
    )
    tf = info_box.text_frame

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f'分析師：{author}'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    run.font.name = 'Calibri'

    p = tf.add_paragraph()
    run = p.add_run()
    run.text = f'日期：{date_str}'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.name = 'Calibri'

    p = tf.add_paragraph()
    run = p.add_run()
    run.text = '本報告僅供內部研究參考，不構成投資建議'
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_cover_slide(
        prs,
        title='台積電 (TSMC)',
        subtitle='深度研究報告 — AI晶片需求持續強勁，上調目標價',
        author='Jamie Chen',
        ticker='2330.TW',
        rating='買入',
        target_price='NT$ 1,200'
    )

    prs.save('封面頁_範例.pptx')
    print('封面頁已生成：封面頁_範例.pptx')
```

### 4.2 數據儀表板頁（KPI卡片+圖表）

```python
"""
範例2：數據儀表板頁
上方4個KPI卡片 + 下方2個圖表
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION


def add_kpi_card(slide, left, top, width, height,
                 label, value, change=None, change_label=None,
                 accent_color='#2B5797'):
    """
    新增KPI指標卡片

    參數：
        label: 指標名稱（如「年化報酬率」）
        value: 指標值（如「23.5%」）
        change: 變化值（如 +2.3 或 -1.5）
        change_label: 變化說明（如「vs 上季」）
    """
    from pptx.oxml.ns import qn
    from lxml import etree

    # 卡片背景
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
    card.line.width = Pt(0.75)

    # 陰影
    spPr = card._element.spPr
    effectLst = etree.SubElement(spPr, qn('a:effectLst'))
    outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', '38100')
    outerShdw.set('dist', '19050')
    outerShdw.set('dir', '5400000')
    outerShdw.set('rotWithShape', '0')
    srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
    srgbClr.set('val', '000000')
    alpha = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha.set('val', '15000')

    # 頂部色條
    color_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, Pt(4)
    )
    color_bar.fill.solid()
    color_bar.fill.fore_color.rgb = RGBColor.from_string(accent_color[1:])
    color_bar.line.fill.background()

    # 指標名稱
    label_box = slide.shapes.add_textbox(
        left + Inches(0.15), top + Inches(0.2),
        width - Inches(0.3), Inches(0.3)
    )
    tf = label_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = label
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.name = 'Calibri'

    # 指標值（大數字）
    value_box = slide.shapes.add_textbox(
        left + Inches(0.15), top + Inches(0.45),
        width - Inches(0.3), Inches(0.6)
    )
    tf = value_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = str(value)
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)
    run.font.name = 'Calibri'

    # 變化值（如果有）
    if change is not None:
        change_box = slide.shapes.add_textbox(
            left + Inches(0.15), top + Inches(1.0),
            width - Inches(0.3), Inches(0.3)
        )
        tf = change_box.text_frame
        p = tf.paragraphs[0]

        # 箭頭符號
        arrow = '\u25B2' if change >= 0 else '\u25BC'  # 上三角/下三角
        color = RGBColor(0x2E, 0x7D, 0x32) if change >= 0 else RGBColor(0xC6, 0x28, 0x28)

        run = p.add_run()
        run.text = f'{arrow} {abs(change):.1f}%'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = color

        if change_label:
            run2 = p.add_run()
            run2.text = f'  {change_label}'
            run2.font.size = Pt(9)
            run2.font.color.rgb = RGBColor(0x99, 0x99, 0x99)


def create_dashboard_slide(prs, title='投資組合績效儀表板'):
    """建立儀表板頁"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # === 頁面標題 ===
    title_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.2), Inches(10), Inches(0.5)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)
    run.font.name = 'Calibri'

    # 分隔線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.7),
        Inches(12.1), Pt(1.5)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
    line.line.fill.background()

    # === 4個KPI卡片 ===
    kpi_data = [
        {'label': '年化報酬率', 'value': '23.5%', 'change': 2.3,
         'change_label': 'vs 基準', 'color': '#2B5797'},
        {'label': '夏普比率', 'value': '1.82', 'change': 0.15,
         'change_label': 'vs 上季', 'color': '#C9A84C'},
        {'label': '最大回撤', 'value': '-12.3%', 'change': -1.2,
         'change_label': 'vs 上季', 'color': '#C62828'},
        {'label': '勝率', 'value': '67.8%', 'change': 3.5,
         'change_label': 'vs 上季', 'color': '#007A7A'},
    ]

    card_width = Inches(2.8)
    card_height = Inches(1.4)
    start_x = Inches(0.6)
    gap = Inches(0.25)

    for i, kpi in enumerate(kpi_data):
        add_kpi_card(
            slide,
            left=start_x + i * (card_width + gap),
            top=Inches(1.0),
            width=card_width,
            height=card_height,
            label=kpi['label'],
            value=kpi['value'],
            change=kpi['change'],
            change_label=kpi['change_label'],
            accent_color=kpi['color']
        )

    # === 左側：淨值曲線圖表 ===
    chart_data = CategoryChartData()
    months = ['1月', '2月', '3月', '4月', '5月', '6月',
              '7月', '8月', '9月', '10月', '11月', '12月']
    chart_data.categories = months
    chart_data.add_series('策略淨值',
        [1.00, 1.03, 1.05, 1.02, 1.08, 1.12, 1.10, 1.15, 1.18, 1.22, 1.20, 1.24])
    chart_data.add_series('基準指數',
        [1.00, 1.01, 1.02, 0.98, 1.01, 1.04, 1.03, 1.06, 1.08, 1.10, 1.09, 1.12])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.6), Inches(2.8), Inches(6.2), Inches(4.2),
        chart_data
    )
    chart = chart_frame.chart
    chart.has_title = True
    chart.chart_title.text_frame.paragraphs[0].text = '淨值曲線比較'
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(11)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(8)

    # 線條配色
    colors_hex = ['2B5797', '999999']
    for i, series in enumerate(chart.series):
        series.format.line.color.rgb = RGBColor.from_string(colors_hex[i])
        series.format.line.width = Pt(2)

    # === 右側：月度報酬率柱狀圖 ===
    bar_data = CategoryChartData()
    bar_data.categories = months
    bar_data.add_series('月報酬率(%)',
        [3.0, 2.0, -2.8, 5.9, 3.7, -1.8, 4.5, 2.6, 3.4, -1.6, 3.3, 0])

    bar_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(7.1), Inches(2.8), Inches(5.6), Inches(4.2),
        bar_data
    )
    bar_chart = bar_frame.chart
    bar_chart.has_title = True
    bar_chart.chart_title.text_frame.paragraphs[0].text = '月度報酬率'
    bar_chart.chart_title.text_frame.paragraphs[0].font.size = Pt(11)
    bar_chart.chart_title.text_frame.paragraphs[0].font.bold = True
    bar_chart.has_legend = False

    # 柱狀圖配色
    series = bar_chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)

    # === 底部來源 ===
    source_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(10), Inches(0.3)
    )
    tf = source_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = 'Source: 回測系統模擬數據 | 過去績效不代表未來表現'
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_dashboard_slide(prs)
    prs.save('儀表板_範例.pptx')
    print('儀表板頁已生成：儀表板_範例.pptx')
```

### 4.3 財務比較表格頁

```python
"""
範例3：財務比較表格頁
投行風格的同業比較表
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def create_financial_comparison_slide(prs):
    """建立財務比較表格頁"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # === 行動標題（McKinsey風格）===
    action_title = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.2), Inches(12), Inches(0.7)
    )
    tf = action_title.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '台積電估值低於歷史平均，相較同業具備明顯折價空間'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)
    run.font.name = 'Calibri'

    # 分隔線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.85),
        Inches(12.1), Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
    line.line.fill.background()

    # === 表格數據 ===
    headers = ['公司', '股票代碼', '市值(億美元)', 'P/E (TTM)',
               'P/E (FY26E)', 'EV/EBITDA', 'ROE (%)',
               '營收增長 (%)', '毛利率 (%)', '評級']

    data = [
        ['台積電', '2330.TW', '8,950', 22.5, 18.3, 14.2, 28.5, 25.3, 55.2, '買入'],
        ['三星電子', '005930.KS', '3,120', 18.7, 15.2, 8.5, 12.3, 8.7, 42.1, '持有'],
        ['英特爾', 'INTC', '1,850', 35.2, 28.5, 18.3, 8.2, -5.2, 38.5, '賣出'],
        ['英偉達', 'NVDA', '28,500', 58.3, 32.1, 42.5, 85.2, 122.5, 72.8, '買入'],
        ['博通', 'AVGO', '7,200', 32.5, 22.8, 22.1, 35.8, 35.2, 68.5, '買入'],
        ['德州儀器', 'TXN', '1,650', 28.3, 24.5, 20.2, 42.1, 2.3, 62.3, '持有'],
    ]

    # 行業中位數
    median_row = ['行業中位數', '—', '—', 30.4, 23.7, 19.3, 27.2, 15.0, 58.8, '—']

    rows = len(data) + 2  # 表頭 + 數據行 + 中位數行
    cols = len(headers)

    table_shape = slide.shapes.add_table(
        rows, cols,
        Inches(0.4), Inches(1.2),
        Inches(12.5), Inches(4.8)
    )
    table = table_shape.table

    # 設定列寬
    col_widths = [Inches(1.4), Inches(1.2), Inches(1.3), Inches(1.0),
                  Inches(1.0), Inches(1.1), Inches(1.0), Inches(1.2),
                  Inches(1.1), Inches(1.2)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = w

    # === 填充表頭 ===
    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x00, 0x33, 0x66)

        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            paragraph.font.size = Pt(9)
            paragraph.font.bold = True
            paragraph.font.name = 'Calibri'
            paragraph.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # === 填充數據 ===
    all_rows = data + [median_row]

    for i, row_data in enumerate(all_rows):
        is_median = (i == len(data))  # 最後一行是中位數
        is_target = (i == 0)  # 第一行是目標公司（台積電）

        for j, value in enumerate(row_data):
            cell = table.cell(i + 1, j)

            # 格式化顯示
            if isinstance(value, float):
                cell.text = f'{value:.1f}'
            else:
                cell.text = str(value)

            # 背景色
            if is_median:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xE8, 0xEE, 0xF5)
            elif is_target:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xF8, 0xE1)
            elif i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF9, 0xF9, 0xF9)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

            # 文字格式
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.name = 'Calibri'

                if is_median:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
                elif is_target:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)
                else:
                    paragraph.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

                # 數值欄位右對齊，文字欄位左對齊
                if j >= 2 and j <= 8 and not isinstance(value, str):
                    paragraph.alignment = PP_ALIGN.RIGHT
                elif j == 9:  # 評級欄
                    paragraph.alignment = PP_ALIGN.CENTER
                    # 評級顏色
                    if value == '買入':
                        paragraph.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)
                        paragraph.font.bold = True
                    elif value == '賣出':
                        paragraph.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
                        paragraph.font.bold = True
                else:
                    paragraph.alignment = PP_ALIGN.LEFT

            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            # 營收增長欄的條件格式（正綠負紅）
            if j == 7 and isinstance(value, float):
                for paragraph in cell.text_frame.paragraphs:
                    if value > 0:
                        paragraph.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)
                    elif value < 0:
                        paragraph.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)

    # === 圖例說明 ===
    legend_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(6.2), Inches(10), Inches(0.5)
    )
    tf = legend_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '■ 黃色背景 = 目標公司  |  ■ 藍色背景 = 行業中位數  |  '
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    run2 = p.add_run()
    run2.text = '綠色 = 正增長  |  紅色 = 負增長'
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # === 來源 ===
    source_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(10), Inches(0.3)
    )
    tf = source_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = 'Source: Bloomberg, 各公司年報 | FY26E 為市場一致預期 | 數據截至 2026年3月'
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_financial_comparison_slide(prs)
    prs.save('財務比較表_範例.pptx')
    print('財務比較表已生成：財務比較表_範例.pptx')
```

### 4.4 技術分析圖表頁

```python
"""
範例4：技術分析圖表頁
使用matplotlib生成K線圖+指標，嵌入PPT
"""
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


def generate_sample_ohlcv(n_days=120, start_price=800):
    """生成模擬K線數據"""
    np.random.seed(42)
    dates = [datetime(2025, 10, 1) + timedelta(days=i) for i in range(n_days)]

    prices = [start_price]
    for _ in range(n_days - 1):
        change = np.random.normal(0.001, 0.02) * prices[-1]
        prices.append(prices[-1] + change)

    opens, highs, lows, closes, volumes = [], [], [], [], []
    for p in prices:
        o = p * (1 + np.random.normal(0, 0.005))
        c = p * (1 + np.random.normal(0, 0.005))
        h = max(o, c) * (1 + abs(np.random.normal(0, 0.008)))
        l = min(o, c) * (1 - abs(np.random.normal(0, 0.008)))
        v = int(np.random.lognormal(15, 0.5))
        opens.append(o)
        highs.append(h)
        lows.append(l)
        closes.append(c)
        volumes.append(v)

    return dates, opens, highs, lows, closes, volumes


def create_candlestick_chart(dates, opens, highs, lows, closes, volumes,
                              title='技術分析', ma_periods=[20, 50]):
    """
    生成投行風格K線圖（含成交量和均線）
    回傳matplotlib figure
    """
    # 設定風格
    plt.rcParams.update({
        'font.family': ['Arial', 'Calibri', 'sans-serif'],
        'axes.facecolor': '#FAFAFA',
        'figure.facecolor': 'none',
        'axes.grid': True,
        'grid.color': '#E8E8E8',
        'grid.linewidth': 0.5,
        'axes.spines.top': False,
        'axes.spines.right': False,
    })

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 6),
                                     gridspec_kw={'height_ratios': [3, 1]},
                                     sharex=True)
    fig.subplots_adjust(hspace=0.05)

    n = len(dates)
    x = np.arange(n)

    # === K線圖 ===
    up = np.array(closes) >= np.array(opens)
    down = ~up

    # 陽線（上漲）
    ax1.bar(x[up], np.array(closes)[up] - np.array(opens)[up],
            bottom=np.array(opens)[up], width=0.6,
            color='#2B5797', edgecolor='#2B5797', linewidth=0.5)
    ax1.vlines(x[up], np.array(lows)[up], np.array(highs)[up],
               color='#2B5797', linewidth=0.5)

    # 陰線（下跌）
    ax1.bar(x[down], np.array(opens)[down] - np.array(closes)[down],
            bottom=np.array(closes)[down], width=0.6,
            color='#C62828', edgecolor='#C62828', linewidth=0.5)
    ax1.vlines(x[down], np.array(lows)[down], np.array(highs)[down],
               color='#C62828', linewidth=0.5)

    # 均線
    ma_colors = {20: '#C9A84C', 50: '#007A7A', 200: '#999999'}
    closes_arr = np.array(closes)
    for period in ma_periods:
        if n >= period:
            ma = np.convolve(closes_arr, np.ones(period)/period, mode='valid')
            ax1.plot(x[period-1:], ma, linewidth=1.2, linestyle='-',
                    color=ma_colors.get(period, '#999'),
                    label=f'MA{period}', alpha=0.9)

    ax1.set_title(title, fontsize=13, fontweight='bold', color='#1A2332', pad=10)
    ax1.legend(loc='upper left', fontsize=8, framealpha=0.9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, p: f'${v:,.0f}'))
    ax1.tick_params(labelsize=8, colors='#666666')

    # === 成交量 ===
    vol_colors = ['#2B5797' if c >= o else '#C62828'
                  for c, o in zip(closes, opens)]
    ax2.bar(x, volumes, width=0.6, color=vol_colors, alpha=0.7)
    ax2.set_ylabel('成交量', fontsize=8, color='#666666')
    ax2.tick_params(labelsize=7, colors='#666666')
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, p: f'{v/1e6:.0f}M' if v >= 1e6 else f'{v/1e3:.0f}K'))

    # X軸日期標籤
    tick_positions = np.linspace(0, n-1, 8).astype(int)
    ax2.set_xticks(tick_positions)
    ax2.set_xticklabels([dates[i].strftime('%m/%d') for i in tick_positions],
                         fontsize=7, rotation=0)

    fig.tight_layout()
    return fig


def create_technical_analysis_slide(prs):
    """建立技術分析頁"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # === 行動標題 ===
    title_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.2), Inches(12), Inches(0.5)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '台積電股價突破MA50壓力線，短期動量轉強，建議逢回佈局'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)

    # 分隔線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.7),
        Inches(12.1), Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
    line.line.fill.background()

    # === K線圖 ===
    dates, opens, highs, lows, closes, volumes = generate_sample_ohlcv()
    fig = create_candlestick_chart(
        dates, opens, highs, lows, closes, volumes,
        title='台積電 (2330.TW) — 日線圖'
    )

    # 嵌入圖表
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=300, transparent=True, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)

    slide.shapes.add_picture(buf, Inches(0.3), Inches(0.9), Inches(9), Inches(5.2))

    # === 右側技術指標面板 ===
    panel_x = Inches(9.5)
    panel_width = Inches(3.5)

    # 面板標題
    panel_title = slide.shapes.add_textbox(
        panel_x, Inches(1.0), panel_width, Inches(0.4)
    )
    tf = panel_title.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '技術指標摘要'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # 指標列表
    indicators = [
        ('RSI (14)', '62.5', '中性偏多', '#2B5797'),
        ('MACD', '正值', '金叉', '#2E7D32'),
        ('KD', '75/68', '高檔鈍化', '#C9A84C'),
        ('布林通道', '中軌上方', '多方格局', '#2E7D32'),
        ('MA20', '$865', '股價在上', '#2E7D32'),
        ('MA50', '$842', '剛突破', '#C9A84C'),
        ('成交量', '放量', '量增價漲', '#2E7D32'),
        ('ATR', '18.5', '波動適中', '#2B5797'),
    ]

    y_start = Inches(1.5)
    for idx, (name, value, signal, color) in enumerate(indicators):
        y = y_start + Inches(idx * 0.55)

        # 指標名稱
        name_box = slide.shapes.add_textbox(panel_x, y, Inches(1.2), Inches(0.25))
        tf = name_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = name
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        run.font.name = 'Calibri'

        # 指標值
        val_box = slide.shapes.add_textbox(
            panel_x + Inches(1.2), y, Inches(1.0), Inches(0.25))
        tf = val_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = value
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)

        # 訊號
        sig_box = slide.shapes.add_textbox(
            panel_x + Inches(2.2), y, Inches(1.3), Inches(0.25))
        tf = sig_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = signal
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(color[1:])

    # === 支撐壓力位 ===
    sr_y = Inches(6.0)
    sr_box = slide.shapes.add_textbox(panel_x, sr_y, panel_width, Inches(1.0))
    tf = sr_box.text_frame

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '關鍵價位'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    levels = [
        ('壓力1', '$920', '#C62828'),
        ('壓力2', '$895', '#C62828'),
        ('支撐1', '$850', '#2E7D32'),
        ('支撐2', '$825', '#2E7D32'),
    ]
    for name, price, color in levels:
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = f'{name}：{price}'
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(color[1:])

    # === 來源 ===
    source_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(10), Inches(0.3)
    )
    tf = source_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = 'Source: TradingView, 自研量化系統 | 技術分析僅供參考，不構成買賣建議'
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_technical_analysis_slide(prs)
    prs.save('技術分析_範例.pptx')
    print('技術分析頁已生成：技術分析_範例.pptx')
```

### 4.5 風險矩陣頁

```python
"""
範例5：風險矩陣頁
影響程度 x 發生機率 的二維矩陣
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


def create_risk_matrix_chart(risks):
    """
    生成風險矩陣散點圖

    risks: [{'name': '風險名稱', 'probability': 0-5, 'impact': 0-5, 'category': '類別'}, ...]
    """
    plt.rcParams.update({
        'font.family': ['Arial', 'Calibri', 'sans-serif'],
        'figure.facecolor': 'none',
        'axes.facecolor': '#FAFAFA',
    })

    fig, ax = plt.subplots(figsize=(8, 6))

    # 背景色塊（低/中/高風險區域）
    # 低風險（綠色）
    ax.fill_between([0, 2], [0, 0], [2, 2], color='#E8F5E9', alpha=0.5)
    ax.fill_between([0, 1], [2, 2], [5, 5], color='#E8F5E9', alpha=0.5)
    ax.fill_between([2, 5], [0, 0], [1, 1], color='#E8F5E9', alpha=0.5)

    # 中風險（黃色）
    ax.fill_between([1, 3], [2, 2], [4, 4], color='#FFF9C4', alpha=0.5)
    ax.fill_between([2, 4], [1, 1], [3, 3], color='#FFF9C4', alpha=0.5)
    ax.fill_between([3, 5], [1, 1], [2, 2], color='#FFF9C4', alpha=0.5)

    # 高風險（紅色）
    ax.fill_between([3, 5], [3, 3], [5, 5], color='#FFEBEE', alpha=0.5)
    ax.fill_between([4, 5], [2, 2], [3, 3], color='#FFEBEE', alpha=0.5)
    ax.fill_between([2, 4], [4, 4], [5, 5], color='#FFEBEE', alpha=0.5)

    # 繪製風險點
    category_colors = {
        '市場風險': '#2B5797',
        '營運風險': '#C9A84C',
        '地緣風險': '#C62828',
        '監管風險': '#007A7A',
        '技術風險': '#6B3FA0',
    }

    for risk in risks:
        color = category_colors.get(risk['category'], '#999999')
        ax.scatter(risk['probability'], risk['impact'],
                  s=200, c=color, edgecolors='white', linewidths=1.5,
                  zorder=5, alpha=0.9)
        ax.annotate(risk['name'],
                   (risk['probability'], risk['impact']),
                   textcoords="offset points", xytext=(8, 8),
                   fontsize=8, color='#333333',
                   bbox=dict(boxstyle='round,pad=0.3',
                           facecolor='white', edgecolor='#D9D9D9',
                           alpha=0.9))

    # 軸設定
    ax.set_xlim(0, 5)
    ax.set_ylim(0, 5)
    ax.set_xlabel('發生機率', fontsize=10, color='#333333', fontweight='bold')
    ax.set_ylabel('影響程度', fontsize=10, color='#333333', fontweight='bold')
    ax.set_title('風險評估矩陣', fontsize=13, fontweight='bold', color='#1A2332', pad=15)

    # 刻度標籤
    labels = ['極低', '低', '中', '高', '極高']
    ax.set_xticks([0.5, 1.5, 2.5, 3.5, 4.5])
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_yticks([0.5, 1.5, 2.5, 3.5, 4.5])
    ax.set_yticklabels(labels, fontsize=8)

    ax.grid(True, color='#E0E0E0', linewidth=0.3)

    # 圖例
    legend_handles = [mpatches.Patch(color=c, label=cat)
                     for cat, c in category_colors.items()]
    ax.legend(handles=legend_handles, loc='lower right', fontsize=8, framealpha=0.9)

    fig.tight_layout()
    return fig


def create_risk_matrix_slide(prs):
    """建立風險矩陣頁"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # === 行動標題 ===
    title_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.2), Inches(12), Inches(0.5)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '地緣政治風險為當前最大隱憂，建議控制倉位並配置避險資產'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)

    # 分隔線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.7),
        Inches(12.1), Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
    line.line.fill.background()

    # === 風險數據 ===
    risks = [
        {'name': '台海局勢升溫', 'probability': 2.5, 'impact': 4.5, 'category': '地緣風險'},
        {'name': '美聯儲意外升息', 'probability': 1.5, 'impact': 4.0, 'category': '市場風險'},
        {'name': '半導體需求下滑', 'probability': 2.0, 'impact': 3.5, 'category': '營運風險'},
        {'name': 'AI監管收緊', 'probability': 3.0, 'impact': 3.0, 'category': '監管風險'},
        {'name': '技術代差縮小', 'probability': 2.0, 'impact': 2.5, 'category': '技術風險'},
        {'name': '日圓大幅升值', 'probability': 3.0, 'impact': 2.0, 'category': '市場風險'},
        {'name': '客戶集中度過高', 'probability': 3.5, 'impact': 3.5, 'category': '營運風險'},
        {'name': '地震等天災', 'probability': 1.0, 'impact': 4.5, 'category': '營運風險'},
    ]

    # === 左側：風險矩陣圖 ===
    fig = create_risk_matrix_chart(risks)
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=300, transparent=True, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)

    slide.shapes.add_picture(buf, Inches(0.3), Inches(1.0), Inches(7.5), Inches(5.5))

    # === 右側：風險詳細列表 ===
    detail_x = Inches(8.2)
    detail_width = Inches(4.8)

    # 面板標題
    panel_title = slide.shapes.add_textbox(
        detail_x, Inches(1.0), detail_width, Inches(0.4)
    )
    tf = panel_title.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '主要風險因素與對策'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # 風險項目（選取前5個高風險）
    sorted_risks = sorted(risks,
                         key=lambda r: r['probability'] * r['impact'],
                         reverse=True)[:5]

    mitigations = {
        '客戶集中度過高': '分散供應鏈，拓展車用/IoT客戶',
        '台海局勢升溫': '配置黃金ETF + 美債作為對沖',
        'AI監管收緊': '關注政策動態，調整AI相關持倉',
        '半導體需求下滑': '監控庫存週期，設置止損位',
        '美聯儲意外升息': '減少高估值成長股配置',
    }

    y = Inches(1.6)
    for i, risk in enumerate(sorted_risks):
        # 風險等級色標
        score = risk['probability'] * risk['impact']
        if score >= 12:
            level_color = '#C62828'
            level_text = '高'
        elif score >= 6:
            level_color = '#F57F17'
            level_text = '中'
        else:
            level_color = '#2E7D32'
            level_text = '低'

        # 色標
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, detail_x, y + Inches(0.05),
            Inches(0.15), Inches(0.15)
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor.from_string(level_color[1:])
        dot.line.fill.background()

        # 風險名稱
        name_box = slide.shapes.add_textbox(
            detail_x + Inches(0.25), y, detail_width - Inches(0.25), Inches(0.25)
        )
        tf = name_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f'{risk["name"]}（{risk["category"]}）'
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)

        # 風險分數
        score_box = slide.shapes.add_textbox(
            detail_x + Inches(0.25), y + Inches(0.25),
            detail_width - Inches(0.25), Inches(0.2)
        )
        tf = score_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f'風險等級：{level_text}（{score:.1f}分）'
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor.from_string(level_color[1:])

        # 對策
        mitigation = mitigations.get(risk['name'], '持續監控')
        mit_box = slide.shapes.add_textbox(
            detail_x + Inches(0.25), y + Inches(0.45),
            detail_width - Inches(0.25), Inches(0.25)
        )
        tf = mit_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f'對策：{mitigation}'
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

        y += Inches(0.95)

    # === 來源 ===
    source_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(10), Inches(0.3)
    )
    tf = source_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = 'Source: 內部風險評估 | 風險等級 = 發生機率 x 影響程度'
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_risk_matrix_slide(prs)
    prs.save('風險矩陣_範例.pptx')
    print('風險矩陣頁已生成：風險矩陣_範例.pptx')
```

### 4.6 投資建議頁

```python
"""
範例6：投資建議頁
結合評級、目標價、建議操作的總結頁
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree


def create_recommendation_slide(prs, company_name='台積電',
                                 ticker='2330.TW',
                                 rating='買入',
                                 current_price='NT$ 880',
                                 target_price='NT$ 1,200',
                                 upside='+36.4%',
                                 time_horizon='12個月'):
    """建立投資建議頁"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # === 行動標題 ===
    title_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.2), Inches(12), Inches(0.5)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f'維持「{rating}」評級 — {company_name}具備長期AI產業鏈核心地位'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)

    # 分隔線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.7),
        Inches(12.1), Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
    line.line.fill.background()

    # === 左側：評級與目標價大卡片 ===
    # 評級卡片背景
    rating_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.1), Inches(4.0), Inches(3.5)
    )
    rating_card.fill.solid()

    # 根據評級設定顏色
    rating_colors = {
        '買入': ('#E8F5E9', '#2E7D32'),
        '增持': ('#E8F5E9', '#558B2F'),
        '持有': ('#FFF9C4', '#F57F17'),
        '減持': ('#FFEBEE', '#E65100'),
        '賣出': ('#FFEBEE', '#C62828'),
    }
    bg_color, text_color = rating_colors.get(rating, ('#F5F5F5', '#333333'))
    rating_card.fill.fore_color.rgb = RGBColor.from_string(bg_color[1:])
    rating_card.line.color.rgb = RGBColor.from_string(text_color[1:])
    rating_card.line.width = Pt(1.5)

    # 投資評級
    rating_label = slide.shapes.add_textbox(
        Inches(0.9), Inches(1.3), Inches(3.5), Inches(0.3)
    )
    tf = rating_label.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = '投資評級'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # 評級值
    rating_value = slide.shapes.add_textbox(
        Inches(0.9), Inches(1.6), Inches(3.5), Inches(0.8)
    )
    tf = rating_value.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = rating
    run.font.size = Pt(42)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(text_color[1:])

    # 目標價
    target_label = slide.shapes.add_textbox(
        Inches(0.9), Inches(2.5), Inches(3.5), Inches(0.3)
    )
    tf = target_label.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = '目標價'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    target_value = slide.shapes.add_textbox(
        Inches(0.9), Inches(2.8), Inches(3.5), Inches(0.6)
    )
    tf = target_value.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = target_price
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x23, 0x32)

    # 上行空間
    upside_box = slide.shapes.add_textbox(
        Inches(0.9), Inches(3.4), Inches(3.5), Inches(0.5)
    )
    tf = upside_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f'當前價 {current_price} → 上行空間 {upside}'
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string(text_color[1:])

    # 時間框架
    time_box = slide.shapes.add_textbox(
        Inches(0.9), Inches(3.9), Inches(3.5), Inches(0.3)
    )
    tf = time_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f'時間框架：{time_horizon}'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    # === 右側：核心投資論點 ===
    thesis_x = Inches(5.2)
    thesis_width = Inches(7.5)

    # 投資論點標題
    thesis_title = slide.shapes.add_textbox(
        thesis_x, Inches(1.1), thesis_width, Inches(0.4)
    )
    tf = thesis_title.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '核心投資論點'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # 投資論點列表
    theses = [
        ('AI晶片需求爆發',
         'CoWoS先進封裝產能滿載至2027年，AI相關營收佔比預計從2025年的15%提升至2026年的25%'),
        ('3nm/2nm技術領先',
         '技術代差維持2年以上優勢，2nm將於2026H2量產，良率已達設計目標'),
        ('全球化佈局降低地緣風險',
         '日本熊本、美國亞利桑那、德國德累斯頓廠分散製造基地'),
        ('定價能力持續提升',
         '先進製程持續漲價5-10%，毛利率有望從55%提升至57-58%'),
    ]

    y = Inches(1.6)
    for i, (title_text, detail_text) in enumerate(theses):
        # 序號圓圈
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            thesis_x, y, Inches(0.3), Inches(0.3)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = RGBColor(0x2B, 0x57, 0x97)
        circle.line.fill.background()

        # 序號文字
        circle_tf = circle.text_frame
        circle_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        circle_run = circle_tf.paragraphs[0].add_run()
        circle_run.text = str(i + 1)
        circle_run.font.size = Pt(10)
        circle_run.font.bold = True
        circle_run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # 論點標題
        point_title = slide.shapes.add_textbox(
            thesis_x + Inches(0.4), y, thesis_width - Inches(0.5), Inches(0.3)
        )
        tf = point_title.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)

        # 論點詳述
        point_detail = slide.shapes.add_textbox(
            thesis_x + Inches(0.4), y + Inches(0.3),
            thesis_width - Inches(0.5), Inches(0.4)
        )
        tf = point_detail.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = detail_text
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

        y += Inches(0.85)

    # === 下方：操作建議面板 ===
    panel_y = Inches(5.0)

    # 面板標題
    op_title = slide.shapes.add_textbox(
        Inches(0.6), panel_y, Inches(12), Inches(0.4)
    )
    tf = op_title.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '操作建議'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # 三個操作建議卡片
    cards = [
        ('積極型投資者', '現價買入，目標持有至NT$1,200\n倉位建議：15-20%',
         '#2E7D32', '#E8F5E9'),
        ('穩健型投資者', '等待回調至MA50（$842）附近分批建倉\n倉位建議：10-15%',
         '#2B5797', '#E3F2FD'),
        ('保守型投資者', '以台積電為核心，搭配半導體ETF分散風險\n倉位建議：5-10%',
         '#C9A84C', '#FFF8E1'),
    ]

    card_width = Inches(3.9)
    card_gap = Inches(0.2)

    for i, (card_title, card_text, accent, bg) in enumerate(cards):
        x = Inches(0.6) + i * (card_width + card_gap)

        # 卡片背景
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x, panel_y + Inches(0.5), card_width, Inches(1.5)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor.from_string(bg[1:])
        card.line.color.rgb = RGBColor.from_string(accent[1:])
        card.line.width = Pt(1)

        # 卡片標題
        ct = slide.shapes.add_textbox(
            x + Inches(0.15), panel_y + Inches(0.6),
            card_width - Inches(0.3), Inches(0.3)
        )
        tf = ct.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = card_title
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(accent[1:])

        # 卡片內容
        cc = slide.shapes.add_textbox(
            x + Inches(0.15), panel_y + Inches(0.95),
            card_width - Inches(0.3), Inches(0.8)
        )
        tf = cc.text_frame
        tf.word_wrap = True
        for line_idx, line_text in enumerate(card_text.split('\n')):
            if line_idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            run = p.add_run()
            run.text = line_text
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # === 免責聲明 ===
    disclaimer = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(12), Inches(0.3)
    )
    tf = disclaimer.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = ('免責聲明：本報告僅供研究參考，不構成投資建議。'
                '投資有風險，過去績效不代表未來表現。請根據自身風險承受能力做出投資決策。')
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run.font.italic = True

    return slide


# === 使用範例 ===
if __name__ == '__main__':
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    create_recommendation_slide(prs)
    prs.save('投資建議_範例.pptx')
    print('投資建議頁已生成：投資建議_範例.pptx')
```

---

## 五、完整整合：一鍵生成完整研究報告

```python
"""
完整整合腳本：一鍵生成投行風格研究報告PPT
將以上所有範例組合成完整的13頁研究報告
"""
from pptx import Presentation
from pptx.util import Inches

def generate_full_report(output_path='投資研究報告.pptx'):
    """
    生成完整的投行風格研究報告
    包含：封面、儀表板、財務比較、技術分析、風險矩陣、投資建議
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 第1頁：封面
    create_cover_slide(
        prs,
        title='台積電 (TSMC)',
        subtitle='深度研究報告 — AI晶片需求持續強勁，上調目標價',
        author='Jamie Chen',
        ticker='2330.TW',
        rating='買入',
        target_price='NT$ 1,200'
    )
    print('  [1/6] 封面頁 ✓')

    # 第2頁：績效儀表板
    create_dashboard_slide(prs, title='台積電 — 關鍵績效指標儀表板')
    print('  [2/6] 儀表板頁 ✓')

    # 第3頁：財務比較表
    create_financial_comparison_slide(prs)
    print('  [3/6] 財務比較頁 ✓')

    # 第4頁：技術分析
    create_technical_analysis_slide(prs)
    print('  [4/6] 技術分析頁 ✓')

    # 第5頁：風險矩陣
    create_risk_matrix_slide(prs)
    print('  [5/6] 風險矩陣頁 ✓')

    # 第6頁：投資建議
    create_recommendation_slide(prs)
    print('  [6/6] 投資建議頁 ✓')

    # 儲存
    prs.save(output_path)
    print(f'\n完整報告已生成：{output_path}')
    return prs


if __name__ == '__main__':
    generate_full_report()
```

---

## 六、進階技巧與最佳實踐

### 6.1 效能優化

```python
# 批量處理時的效能建議

# 1. 重複使用Presentation物件，避免頻繁建立
# 2. matplotlib圖表用BytesIO避免磁碟IO
# 3. 大量圖片使用適當的DPI（螢幕150-200，列印300）
# 4. 使用lru_cache快取重複計算

from functools import lru_cache

@lru_cache(maxsize=32)
def get_color_rgb(hex_color):
    """快取顏色轉換"""
    return RGBColor.from_string(hex_color.lstrip('#'))
```

### 6.2 模板管理

```python
class PPTTemplate:
    """
    PPT模板管理器
    統一管理配色、字體、版面等設定
    """

    # === 配色方案 ===
    COLORS = {
        'primary':     '#003366',
        'secondary':   '#2B5797',
        'accent':      '#C9A84C',
        'positive':    '#2E7D32',
        'negative':    '#C62828',
        'warning':     '#F57F17',
        'text_dark':   '#1A2332',
        'text_medium': '#333333',
        'text_light':  '#666666',
        'text_muted':  '#999999',
        'border':      '#D9D9D9',
        'bg_light':    '#F5F5F5',
        'white':       '#FFFFFF',
    }

    # === 字體 ===
    FONT_TITLE = 'Calibri'
    FONT_BODY = 'Calibri'
    FONT_DATA = 'Calibri'
    FONT_CN = '微軟雅黑'  # 中文備用

    # === 字號 ===
    FONT_SIZE = {
        'page_title': Pt(20),
        'action_title': Pt(16),
        'section_title': Pt(14),
        'body': Pt(11),
        'small': Pt(9),
        'footnote': Pt(7),
        'kpi_number': Pt(36),
    }

    # === 版面 ===
    MARGIN = {
        'left': Inches(0.6),
        'right': Inches(0.6),
        'top': Inches(0.5),
        'bottom': Inches(0.4),
    }

    @classmethod
    def get_color(cls, name):
        """取得顏色的RGBColor物件"""
        hex_val = cls.COLORS.get(name, '#000000')
        return RGBColor.from_string(hex_val.lstrip('#'))
```

### 6.3 常見問題與解決方案

| 問題 | 原因 | 解決方案 |
|------|------|---------|
| 中文顯示為方框 | 字體不支援中文 | 使用微軟雅黑或思源黑體 |
| 圖表模糊 | DPI太低 | matplotlib使用dpi=300 |
| 檔案太大 | 圖片未壓縮 | 控制DPI，使用JPEG代替PNG |
| 表格超出頁面 | 列寬計算錯誤 | 使用SlideGrid類計算 |
| 漸變填充無效 | API不支援 | 使用XML注入方式 |
| 投影片順序錯誤 | add_slide只能追加 | 按順序建立，或用XML重排 |
| 字體嵌入問題 | 跨平台字體不同 | 使用系統通用字體或嵌入字體 |

---

## 七、參考資源

### 設計參考
- Goldman Sachs Global Investment Research 報告風格
- Morgan Stanley Research 報告格式
- McKinsey Presentation Framework（金字塔原則 + MECE）
- Edward Tufte《The Visual Display of Quantitative Information》

### 技術文件
- python-pptx 官方文件：https://python-pptx.readthedocs.io/
- matplotlib 官方文件：https://matplotlib.org/stable/
- Open XML SDK 規範（理解底層XML結構）

### 投行字體使用慣例
- Goldman Sachs / Morgan Stanley：Times New Roman
- J.P. Morgan / Bank of America：Arial
- Citi / Deutsche Bank：Calibri
- UBS / Barclays：Helvetica
- McKinsey：Arial（正文）+ Georgia（標題）

### 配色資源
- BrandColorCode.com — 品牌色號查詢
- ColorBrewer2.org — 數據可視化配色
- Coolors.co — 配色方案生成器

---

> 本指南整合了頂級投行的設計規範、python-pptx的高級技巧、以及完整的可運行代碼範例。
> 所有代碼範例均可獨立運行，也可組合使用生成完整的研究報告。
