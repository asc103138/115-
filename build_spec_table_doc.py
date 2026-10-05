import os
import subprocess
import docx
from docx import Document
from docx.shared import Mm, Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

doc = Document()

# 設定 A4 橫向 (Landscape)
sec = doc.sections[0]
sec.page_width = Mm(297)
sec.page_height = Mm(210)
sec.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
sec.left_margin = Mm(15)
sec.right_margin = Mm(15)
sec.top_margin = Mm(15)
sec.bottom_margin = Mm(15)

ns = nsdecls("w")

# 設定 Normal 樣式字型為標楷體
style_normal = doc.styles["Normal"]
style_normal.font.name = "標楷體"
rPr_normal = style_normal.element.get_or_add_rPr()
rFonts_normal = parse_xml(f'<w:rFonts {ns} w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia"/>')
rPr_normal.append(rFonts_normal)

def add_kai_run(p, text, size=11, bold=False, color=None):
    r = p.add_run(text)
    r.font.name = "標楷體"
    r.font.size = Pt(size)
    r.bold = bold
    if color:
        r.font.color.rgb = color
    rPr = r._element.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {ns} w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia"/>')
    rPr.append(rFonts)
    return r

def add_p(text="", size=11, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=3, line_spacing=15):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(line_spacing)
    if text:
        add_kai_run(p, text, size=size, bold=bold)
    return p

def set_cell_border(cell, **kwargs):
    """
    kwargs: top, bottom, left, right
    values: {"sz": 4, "val": "single", "color": "000000"}
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {ns}/>')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'<w:{edge} {ns} w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", 4)}" w:space="0" w:color="{edge_data.get("color", "000000")}"/>'
        else:
            tag = f'<w:{edge} {ns} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex="F2F2F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {ns} w:val="clear" w:color="auto" w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {ns}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

# ==================== 標題與基本資料 ====================
add_p("臺中市梧棲區中正國民小學 115 學年度第一學期四年級國語科第一次定期成績評量", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_p("【 命題雙向細目表與評量分析手冊 】", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

p_info = add_p(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
add_kai_run(p_info, "評量年級：四年級　　命題範圍：第 1 課至第 6 課、統整活動（一）　　命題者：四年級命題小組　　審題者：四年級教學研究會", size=10.5)

# ==================== 壹、評量基本架構與配分統計表 ====================
add_p("壹、大題架構與配分統計表", size=12, bold=True, space_before=4, space_after=4)

table1 = doc.add_table(rows=1, cols=6)
table1.alignment = WD_TABLE_ALIGNMENT.CENTER
headers1 = ["大題編號", "大題名稱", "題數／格數", "配分", "佔比", "評量重點與題型說明"]
widths1 = [Mm(20), Mm(38), Mm(25), Mm(18), Mm(18), Mm(148)]

hdr_cells1 = table1.rows[0].cells
for i, name in enumerate(headers1):
    hdr_cells1[i].width = widths1[i]
    set_cell_border(hdr_cells1[i])
    set_cell_shading(hdr_cells1[i], "E8EEF5")
    set_cell_margins(hdr_cells1[i], top=120, bottom=120, left=120, right=120)
    p = hdr_cells1[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_kai_run(p, name, size=10, bold=True)

data1 = [
    ("一", "寫出國字和注音", "28 格", "28 分", "28%", "融合四上 1~6 課情境短文（14格考國字＋14格考注音），雙方框標準題型"),
    ("二", "改錯字", "6 題", "12 分", "12%", "生活情境短文（挫折調適），測驗形近字、音近字辨析與正字書寫"),
    ("三", "選擇題測驗", "10 題", "20 分", "20%", "單選題（四選一），涵蓋擬人修辭、書信格式、讀報策略、複句轉換、SEL態度"),
    ("四", "成語填空", "5 題", "10 分", "10%", "六選五代號填空，測驗常用成語（垂涎三尺、大排長龍、未雨綢繆等）之語境運用"),
    ("五", "句型練習", "4 題", "14 分", "14%", "（一）短語仿寫 2 題（6分）＋（二）複句造句 2 題（8分，遞進與讓步假設）"),
    ("六", "閱讀理解", "4 題", "16 分", "16%", "科普說明文〈火星探測車毅力號〉，測驗訊息提取、因果推論與人生省思"),
    ("合計", "全卷共 6 大題", "57 評量點", "100 分", "100%", "基礎概念（記＋理）64% ｜ 高階思維 36% ｜ 難易度比約 3 : 5 : 2")
]

for row_data in data1:
    row = table1.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.width = widths1[i]
        set_cell_border(cell)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        if i in [0, 2, 3, 4]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        is_bold = (row_data[0] == "合計")
        if is_bold:
            set_cell_shading(cell, "F2F2F2")
        add_kai_run(p, val, size=9.5, bold=is_bold)

# ==================== 貳、認知歷程向度與 PIRLS 閱讀歷程統計 ====================
doc.add_page_break()
add_p("貳、認知歷程向度（Bloom）與 PIRLS 閱讀歷程統計分析", size=12, bold=True, space_before=4, space_after=4)

table2 = doc.add_table(rows=1, cols=5)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
headers2 = ["認知向度層次（Bloom）", "評量題數／格數", "配分", "百分比", "涵蓋試題大題與評量重點說明"]
widths2 = [Mm(45), Mm(28), Mm(20), Mm(20), Mm(154)]

for i, name in enumerate(headers2):
    cell = table2.rows[0].cells[i]
    cell.width = widths2[i]
    set_cell_border(cell)
    set_cell_shading(cell, "E8EEF5")
    set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_kai_run(p, name, size=10, bold=True)

data2 = [
    ("1. 記憶 (Remember)", "32 評量點", "38 分", "38%", "一大題（國字14＋注音14共28分）、三大題 Q2(書信)、Q6(讀報)、六大題 Q1/Q2(事實提取)"),
    ("2. 理解 (Understand)", "13 評量點", "26 分", "26%", "二大題（改錯6題共12分）、三大題 Q3(主旨)/Q4(細節)/Q7(段落)/Q8(複句)、四大題 Q1/Q2/Q4"),
    ("3. 應用 (Apply)", "5 評量點", "14 分", "14%", "四大題 Q3/Q5(成語語用4分)、五大題短語仿寫(6分)、三大題 Q5(譬喻效果遷移2分)、四大題 Q5(2分)"),
    ("4. 分析 (Analyze)", "2 評量點", "4 分", "4%", "三大題 Q1(擬人修辭效果分析2分)、三大題 Q9(詞語反義關係辨析2分)"),
    ("5. 評鑑 (Evaluate)", "2 評量點", "6 分", "6%", "三大題 Q10(SEL情緒調適策略評鑑2分)、六大題 Q4(科學精神與人生哲理評鑑4分)"),
    ("6. 創造 (Create)", "3 評量點", "12 分", "12%", "五大題造句 2 題（遞進與讓步複句8分）、六大題 Q3(問題排除策略推論與情境結合4分)"),
    ("合計", "57 評量點", "100 分", "100%", "布魯姆六大認知層次全覆蓋，高階思維（應用/分析/評鑑/創造）佔 36 分")
]

for row_data in data2:
    row = table2.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.width = widths2[i]
        set_cell_border(cell)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        if i in [1, 2, 3]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        is_bold = (row_data[0] == "合計")
        if is_bold:
            set_cell_shading(cell, "F2F2F2")
        add_kai_run(p, val, size=9.5, bold=is_bold)

add_p("", size=6)
add_p("【 PIRLS 四大閱讀歷程佔比（針對選擇題與閱讀理解題群，共 14 題 36 分）】", size=10.5, bold=True, space_after=3)

table_pirls = doc.add_table(rows=1, cols=4)
table_pirls.alignment = WD_TABLE_ALIGNMENT.CENTER
headers_pirls = ["PIRLS 閱讀歷程層次", "配分", "佔比 (%)", "涵蓋試題題號與評量核心"]
widths_pirls = [Mm(55), Mm(25), Mm(25), Mm(162)]

for i, name in enumerate(headers_pirls):
    cell = table_pirls.rows[0].cells[i]
    cell.width = widths_pirls[i]
    set_cell_border(cell)
    set_cell_shading(cell, "F0F4F8")
    set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_kai_run(p, name, size=9.5, bold=True)

data_pirls = [
    ("R1 直接提取 (Retrieve)", "14 分", "38.9%", "三大題 Q2(書信結尾)、Q4(飛行夢細節)、Q6(讀報重點)｜六大題 Q1(毅力號任務)、Q2(火星環境)"),
    ("R2 直接推論 (Infer)", "8 分", "22.2%", "三大題 Q5(譬喻修辭推論)、Q8(即使也近義轉換)｜六大題 Q3(排除碎石方法推論)"),
    ("R3 詮釋整合 (Interpret)", "4 分", "11.1%", "三大題 Q3(鏡頭下的家鄉主旨整合)、三大題 Q10(挫折因應之 SEL 心理健康整合)"),
    ("R4 比較評估 (Evaluate)", "10 分", "27.8%", "三大題 Q1(吐著唾沫表達效果評估)、Q7(段落概念評估)、Q9(詞語關係對比)｜六大題 Q4(人生態度評鑑)"),
    ("合計", "36 分", "100.0%", "兼顧客觀事實提取 (38.9%) 與深層推論評估 (61.1%)，符合素養導向評量標準")
]

for row_data in data_pirls:
    row = table_pirls.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.width = widths_pirls[i]
        set_cell_border(cell)
        set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
        p = cell.paragraphs[0]
        if i in [1, 2]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        is_bold = (row_data[0] == "合計")
        if is_bold:
            set_cell_shading(cell, "F2F2F2")
        add_kai_run(p, val, size=9, bold=is_bold)

# ==================== 參、逐題命題雙向細目表 ====================
doc.add_page_break()
add_p("參、全卷逐題命題雙向細目表（Item-by-Item Specification Matrix）", size=12, bold=True, space_before=4, space_after=4)

table3 = doc.add_table(rows=1, cols=8)
table3.alignment = WD_TABLE_ALIGNMENT.CENTER
headers3 = ["題號", "大題分項", "評量目標字詞／題目重點", "教材出處", "認知向度", "閱讀歷程", "難易度", "配分"]
widths3 = [Mm(14), Mm(24), Mm(75), Mm(40), Mm(26), Mm(38), Mm(16), Mm(14)]

for i, name in enumerate(headers3):
    cell = table3.rows[0].cells[i]
    cell.width = widths3[i]
    set_cell_border(cell)
    set_cell_shading(cell, "E8EEF5")
    set_cell_margins(cell, top=100, bottom=100, left=80, right=80)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_kai_run(p, name, size=9.5, bold=True)

items3 = [
    # 一大題
    ("一-1", "國字注音", "彷（ㄈㄤˇ彿）— 彳部，易誤作「仿」", "L01 美麗島", "記憶", "字形書寫", "易", "1"),
    ("一-2", "國字注音", "鷹（【ㄧㄥ】）— 鳥部，零聲母一聲", "L01 美麗島", "記憶", "字音認讀", "易", "1"),
    ("一-3", "國字注音", "衝（ㄔㄨㄥ刺）— 行部，重字筆順", "L04 飛行夢", "記憶", "字形書寫", "中", "1"),
    ("一-4", "國字注音", "竟（【ㄐㄧㄥˋ】然）— 結合韻ㄧㄥ四聲", "L04 飛行夢", "記憶", "字音認讀", "易", "1"),
    ("一-5", "國字注音", "訝（驚ㄧㄚˋ）— 言部，右牙勿多撇", "L03 鏡頭下的家鄉", "記憶", "字形書寫", "易", "1"),
    ("一-6", "國字注音", "恐（【ㄎㄨㄥˇ】懼）— 雙唇音ㄎ三聲", "L04 飛行夢", "記憶", "字音認讀", "中", "1"),
    ("一-7", "國字注音", "樹（ㄕㄨˋ下）— 木＋壴＋寸結構平衡", "L01 美麗島", "記憶", "字形書寫", "易", "1"),
    ("一-8", "國字注音", "蕉（香【ㄐㄧㄠ】）— 艸部，結合韻ㄧㄠ一聲", "L01 美麗島", "記憶", "字音認讀", "易", "1"),
    ("一-9", "國字注音", "狀（見ㄓㄨㄤˋ）— 爿部，右犬勿漏點", "L06 又遠又近的月亮", "記憶", "字形書寫", "中", "1"),
    ("一-10", "國字注音", "隊（成群結【ㄉㄨㄟˋ】）— 結合韻ㄨㄟ四聲", "L01 美麗島", "記憶", "字音認讀", "易", "1"),
    ("一-11", "國字注音", "寧（ㄋㄧㄥˊ靜）— 宀部，心皿結構繁密", "L05 月光下", "記憶", "字形書寫", "難", "1"),
    ("一-12", "國字注音", "控（【ㄎㄨㄥˋ】制）— 結合韻ㄨㄥ四聲", "L04 飛行夢", "記憶", "字音認讀", "易", "1"),
    ("一-13", "國字注音", "制（控ㄓˋ）— 刀部，左製易多橫", "L04 飛行夢", "記憶", "字形書寫", "中", "1"),
    ("一-14", "國字注音", "碑（石【ㄅㄟ】）— 雙唇音ㄅ一聲", "L04 飛行夢", "記憶", "字音認讀", "易", "1"),
    ("一-15", "國字注音", "巡（ㄒㄩㄣˊ視）— 巛部，辶字旁筆順", "L05 月光下", "記憶", "字形書寫", "中", "1"),
    ("一-16", "國字注音", "甜（【ㄊㄧㄢˊ】蜜）— 結合韻ㄧㄢ二聲", "L01 美麗島", "記憶", "字音認讀", "易", "1"),
    ("一-17", "國字注音", "航（【ㄏㄤˊ】程）— 舌根音ㄏ二聲", "L04 飛行夢", "記憶", "字音認讀", "易", "1"),
    ("一-18", "國字注音", "糖（【ㄊㄤˊ】廠）— 米部，唐字結構部件", "L01 美麗島", "記憶", "字形書寫", "易", "1"),
    ("一-19", "國字注音", "榮（光【ㄖㄨㄥˊ】）— 捲舌音ㄖ結合韻ㄨㄥ", "L02 請到我的家鄉來", "記憶", "字音認讀", "中", "1"),
    ("一-20", "國字注音", "衛（守【ㄨㄟˋ】）— 行部，韋字部件平衡", "L01 美麗島", "記憶", "字形書寫", "中", "1"),
    ("一-21", "國字注音", "灣（海【ㄨㄢ】）— 介音ㄨ結合韻ㄢ", "L01 美麗島", "記憶", "字音認讀", "易", "1"),
    ("一-22", "國字注音", "嶼（島ㄩˇ）— 山部，與字筆畫結構", "L01 美麗島", "記憶", "字形書寫", "難", "1"),
    ("一-23", "國字注音", "翔（翱【ㄒㄧㄤˊ】）— 羽部，羊＋羽部件", "L04 飛行夢", "記憶", "字形書寫", "中", "1"),
    ("一-24", "國字注音", "研（【ㄧㄢˊ】究）— 石部，開字部件平衡", "L04 飛行夢", "記憶", "字音認讀", "易", "1"),
    ("一-25", "國字注音", "造（改ㄗㄠˋ）— 辶部，告字部件", "L04 飛行夢", "記憶", "字形書寫", "易", "1"),
    ("一-26", "國字注音", "篇（每一【ㄆㄧㄢ】）— 竹部結合韻ㄧㄢ", "L02 請到我的家鄉來", "記憶", "字音認讀", "易", "1"),
    ("一-27", "國字注音", "盤（銀ㄆㄢˊ）— 皿部，舟＋殳＋皿", "L05 月光下", "記憶", "字形書寫", "中", "1"),
    ("一-28", "國字注音", "耀（閃【ㄧㄠˋ】）— 羽部結合韻ㄧㄠ四聲", "L03 鏡頭下的家鄉", "記憶", "字音認讀", "易", "1"),
    # 二大題 改錯字
    ("二-1", "改錯字", "自【則】→ 責（形近貝部責 vs 刀部則）", "L04 語文焦點", "理解", "錯別辨析", "易", "2"),
    ("二-2", "改錯字", "滿【惱】子 → 腦（音近肉部腦 vs 心部惱）", "L04 課文生字", "理解", "錯別辨析", "中", "2"),
    ("二-3", "改錯字", "挫【拆】→ 折（形近手部折 vs 拆多一點）", "L04 課文生字", "理解", "錯別辨析", "易", "2"),
    ("二-4", "改錯字", "彷【服】→ 彿（音近彳部彿 vs 月部服）", "L01 課文生字", "理解", "錯別辨析", "中", "2"),
    ("二-5", "改錯字", "【距】絕 → 拒（音近手部拒 vs 足部距）", "L03 課文生字", "理解", "錯別辨析", "易", "2"),
    ("二-6", "改錯字", "【空】制 → 控（音近手部控 vs 穴部空）", "L04 課文生字", "理解", "錯別辨析", "中", "2"),
    # 三大題 選擇題
    ("三-1", "選擇題", "〈美麗島〉「吐著唾沫」寫作效果（擬人）", "L01 美麗島", "分析", "R4 比較評估", "中", "2"),
    ("三-2", "選擇題", "〈請到我家鄉來〉信件結尾格式（祝福/署名/日期）", "L02 請到我家鄉來", "記憶", "R1 直接提取", "易", "2"),
    ("三-3", "選擇題", "〈鏡頭下的家鄉〉人物照片主要傳達意涵（生命力）", "L03 鏡頭下的家鄉", "理解", "R3 詮釋整合", "中", "2"),
    ("三-4", "選擇題", "〈飛行夢〉萊特兄弟面對挫折不放棄的堅持精神", "L04 飛行夢", "理解", "R1 直接提取", "易", "2"),
    ("三-5", "選擇題", "〈月光下〉「呼作白玉盤」譬喻手法辨析", "L05 月光下", "應用", "R2 直接推論", "難", "2"),
    ("三-6", "選擇題", "讀報策略：短時間掌握新聞核心優先讀標題", "L06 讀報策略", "理解", "R1 直接提取", "易", "2"),
    ("三-7", "選擇題", "自然段（換行退兩格）與意義段之本質概念區辨", "統整活動一", "理解", "R4 比較評估", "中", "2"),
    ("三-8", "選擇題", "「即使……也……」讓步句型近義轉換（就算……還是）", "統整活動一", "理解", "R2 直接推論", "中", "2"),
    ("三-9", "選擇題", "成語語詞關係（成群結隊／形單影隻 為相反詞）", "統整活動一", "分析", "R4 比較評估", "難", "2"),
    ("三-10", "選擇題", "SEL 情緒調適：面對失敗向信任師長傾訴求助", "跨課情意統整", "評鑑", "R3 詮釋整合", "易", "2"),
    # 四大題 成語填空
    ("四-1", "成語填空", "紅燒肉香氣四溢 → (2) 垂涎三尺", "統整活動一", "理解", "語境選填", "易", "2"),
    ("四-2", "成語填空", "售票口一票難求 → (6) 大排長龍", "統整活動一", "理解", "語境選填", "易", "2"),
    ("四-3", "成語填空", "颱風前準備防災 → (4) 未雨綢繆", "統整活動一", "應用", "語境選填", "中", "2"),
    ("四-4", "成語填空", "出國消息令人意外 → (5) 大吃一驚", "統整活動一", "理解", "語境選填", "易", "2"),
    ("四-5", "成語填空", "校慶活動節目精彩 → (1) 多采多姿", "統整活動一", "應用", "語境選填", "中", "2"),
    # 五大題 句型練習
    ("五-1", "短語仿寫", "（ 又遠又近 ）的（ 月亮 ）【又A又B的C】", "L05 課文短語", "應用", "語法仿寫", "易", "3"),
    ("五-2", "短語仿寫", "（ 陽光 ）如（ 金粉 ）般（ 落下 ）【A如B般C】", "L01 課文短語", "應用", "譬喻動態", "中", "3"),
    ("五-3", "造句練習", "不只……而且…… ──（遞進複句）", "統整活動一", "創造", "複句邏輯", "中", "4"),
    ("五-4", "造句練習", "即使……也…… ──（讓步假設複句）", "統整活動一", "創造", "複句邏輯", "中", "4"),
    # 六大題 閱讀理解
    ("六-1", "閱讀理解", "毅力號火星探測車主要任務（生命遺跡與採樣）", "科普跨領域", "記憶", "R1 直接提取", "易", "4"),
    ("六-2", "閱讀理解", "火星極端環境特點（溫差極大且常有沙塵暴）", "科普跨領域", "記憶", "R1 直接提取", "易", "4"),
    ("六-3", "閱讀理解", "科學家排除採樣管碎石方法（分析模擬震動指令）", "科普跨領域", "理解", "R2 直接推論", "中", "4"),
    ("六-4", "閱讀理解", "本文主要傳達道理（面對困難冷靜分析堅持不懈）", "科普跨領域", "評鑑", "R4 比較評估", "中", "4"),
]

for row_data in items3:
    row = table3.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.width = widths3[i]
        set_cell_border(cell)
        set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
        p = cell.paragraphs[0]
        if i in [0, 1, 4, 5, 6, 7]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_kai_run(p, val, size=8.5)

# ==================== 肆、標準答案與評分規準總表 ====================
doc.add_page_break()
add_p("肆、標準解答總覽與給分扣分規準", size=12, bold=True, space_before=4, space_after=4)

table4 = doc.add_table(rows=1, cols=4)
table4.alignment = WD_TABLE_ALIGNMENT.CENTER
headers4 = ["大題名稱", "題號／項目", "標準答案與參考範例", "扣分原則與給分注意事項"]
widths4 = [Mm(30), Mm(30), Mm(110), Mm(97)]

for i, name in enumerate(headers4):
    cell = table4.rows[0].cells[i]
    cell.width = widths4[i]
    set_cell_border(cell)
    set_cell_shading(cell, "E8EEF5")
    set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_kai_run(p, name, size=10, bold=True)

data4 = [
    ("一、寫出國字和注音", "1~28 格\n(共28分)", "國字：彷、衝、訝、樹、狀、寧、制、巡、糖、衛、嶼、翔、造、盤\n注音：ㄧㄥ、ㄐㄧㄥˋ、ㄎㄨㄥˇ、ㄐㄧㄠ、ㄉㄨㄟˋ、ㄎㄨㄥˋ、ㄅㄟ、ㄊㄧㄢˊ、ㄏㄤˊ、ㄖㄨㄥˊ、ㄨㄢ、ㄧㄢˊ、ㄆㄧㄢ、ㄧㄠˋ", "• 每格 1 分。\n• 國字多一筆少一畫不給分。\n• 注音符號、介音或聲調錯誤不給分。"),
    ("二、改錯字", "1~6 題\n(共12分)", "1. 責（則） 2. 腦（惱） 3. 折（拆）\n4. 彿（服） 5. 拒（距） 6. 控（空）", "• 每題 2 分（圈出錯字得 1 分，寫對正字得 1 分）。\n• 填寫順序調換不扣分；正字寫錯別字扣 1 分。"),
    ("三、選擇題", "1~10 題\n(共20分)", "1. (1)  2. (2)  3. (3)  4. (3)  5. (1)\n6. (3)  7. (1)  8. (2)  9. (3) 10. (3)", "• 每題 2 分，單選題，答錯不倒扣。"),
    ("四、成語填空", "1~5 題\n(共10分)", "1. (2) 垂涎三尺  2. (6) 大排長龍  3. (4) 未雨綢繆\n4. (5) 大吃一驚  5. (1) 多采多姿", "• 每題 2 分，須填寫代號。\n• 若填寫國字全名且完全正確，依研究會決議給分或扣1分。"),
    ("五、句型練習", "（一）短語仿寫\n(共6分)", "1. 例：（ 又香又甜 ）的（ 西瓜 ）【又A又B的C】\n2. 例：（ 白雪 ）如（ 鵝毛 ）般（ 飄落 ）【A如B般C】", "• 每題 3 分。\n• 詞性結構相符且語意通順給 3 分。\n• 結構不完整或詞性不合扣 1~2 分；錯別字每字扣 1 分。"),
    ("五、句型練習", "（二）造句練習\n(共8分)", "1. 例：他不但成績優異，而且熱心助人。\n2. 例：即使風雨再大，我們也堅持準時到校。", "• 每題 4 分。\n• 關聯複句邏輯明確且標點符號正確給 4 分。\n• 複句誤用（如讓步寫成因果）扣 2 分；錯別字每字扣 1 分。"),
    ("六、閱讀理解", "1~4 題\n(共16分)", "1. (3) 古代生命遺跡並收集樣本\n2. (2) 溫差極大且常有沙塵暴\n3. (3) 分析模擬並發送震動指令\n4. (2) 面對困難冷靜分析堅持不懈", "• 每題 4 分，單選題，答錯不倒扣。")
]

for row_data in data4:
    row = table4.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.width = widths4[i]
        set_cell_border(cell)
        set_cell_margins(cell, top=90, bottom=90, left=100, right=100)
        p = cell.paragraphs[0]
        if i == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_kai_run(p, val, size=9)

out_docx = "115四上國語期中評量試卷_雙向細目表.docx"
doc.save(out_docx)
print(f"Saved {out_docx}")
