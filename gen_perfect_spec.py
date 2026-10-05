import os
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import win32com.client
import pymupdf

def generate_perfect_spec():
    src_docx = r"d:\opencode\115-\template_spec.docx"
    doc = Document(src_docx)
    ns = nsdecls("w")
    
    # Page setup: Standard A4 portrait, margins 20mm
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(18)
    sec.bottom_margin = Mm(18)
    sec.left_margin = Mm(20)
    sec.right_margin = Mm(20)
    
    # Normal style
    style = doc.styles["Normal"]
    style.font.name = "標楷體"
    style.font.size = Pt(11)
    style._element.get_or_add_rPr().append(
        parse_xml(f'<w:rFonts {ns} w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia"/>')
    )
    
    def set_run_kai(r, font_size=11, bold=False):
        r.font.name = "標楷體"
        r.font.size = Pt(font_size)
        r.bold = bold
        r._element.get_or_add_rPr().append(
            parse_xml(f'<w:rFonts {ns} w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia"/>')
        )

    def set_cell_text(cell, text, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.text = text
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.alignment = align
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                set_run_kai(r, font_size=size, bold=bold)

    # 1. Title paragraphs
    # p0: 梧棲區中正國小115學年度第一學期第一次評量試題分析
    # p1: 雙向細目表
    # p2: 科目名稱：國語            命題教師：四年級命題小組            使用年級：四年級
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "科目名稱" in txt:
            p.text = "科目名稱：國語            命題教師：四年級命題小組            使用年級：四年級"
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_kai(r, font_size=11, bold=True)
        elif "雙向細目表" in txt and len(txt) < 10:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_kai(r, font_size=16, bold=True)
        elif "試題分析" in txt:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                set_run_kai(r, font_size=16, bold=True)
        elif "一、命題範圍" in txt:
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_kai(r, font_size=12, bold=True)
        elif "二、測驗的題型" in txt:
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_kai(r, font_size=12, bold=True)
        elif "三、試題分析雙向" in txt:
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_kai(r, font_size=12, bold=True)
        elif "【說明】" in txt:
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                set_run_kai(r, font_size=9.5, bold=True)
        elif any(k in txt for k in ["理想配分=", "「實際配分」", "命題原則：", "表格不夠時", "請命題者依", "6(2)表示", "表格請自行"]):
            p.paragraph_format.space_before = Pt(0.5)
            p.paragraph_format.space_after = Pt(0.5)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                set_run_kai(r, font_size=9)

    # 2. Table 1
    t1 = doc.tables[0]
    
    # We will remove extra unused rows (rows 10, 11, 12, 13) so that Table 1 has 10 data rows + 1 total row = 11 rows!
    # Wait, in t1:
    # Row 0: Header
    # Rows 1~3: Unit 1 (3 lessons)
    # Rows 4~6: Unit 2 (3 lessons)
    # Rows 7~9: Integration & Review (3 activities)
    # Row 10~13: Extra unused rows in template
    # Row 14: Total row
    
    # Let's delete rows 13, 12, 11, 10
    for _ in range(4):
        row_to_del = t1.rows[10]
        t1._tbl.remove(row_to_del._tr)
        
    # Now t1 has 11 rows: row 0 (header), rows 1~9 (content), row 10 (total)
    
    # Fill Block 1 (Rows 1~3, Unit 1)
    set_cell_text(t1.cell(1, 0), "第一單元：\n好山好水\n好故鄉", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    t1_block1 = [
        ("第 1 課 美麗島（詩歌節奏、動態描寫與寶島情感）", "5", "11 ％", "11 ％"),
        ("第 2 課 請到我的家鄉來（書信閱讀、景點特色與互動）", "6", "13 ％", "13 ％"),
        ("第 3 課 鏡頭下的家鄉（攝影細節、家鄉人物與生活）", "6", "13 ％", "13 ％"),
    ]
    for r_idx, (act, sec_cnt, idl, act_p) in enumerate(t1_block1, start=1):
        set_cell_text(t1.cell(r_idx, 1), act, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(t1.cell(r_idx, 2), sec_cnt, size=10)
        set_cell_text(t1.cell(r_idx, 3), idl, size=10)
        set_cell_text(t1.cell(r_idx, 4), act_p, size=10)

    # Fill Block 2 (Rows 4~6, Unit 2)
    set_cell_text(t1.cell(4, 0), "第二單元：\n天空的奇想", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    t1_block2 = [
        ("第 4 課 飛行夢（萊特兄弟、面對挫折與堅持夢想）", "6", "13 ％", "13 ％"),
        ("第 5 課 月光下（河邊賞月、美景四字語詞與比喻）", "6", "13 ％", "13 ％"),
        ("第 6 課 又遠又近的月亮（登月報導、六何要素與科普）", "6", "14 ％", "14 ％"),
    ]
    for r_idx, (act, sec_cnt, idl, act_p) in enumerate(t1_block2, start=4):
        set_cell_text(t1.cell(r_idx, 1), act, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(t1.cell(r_idx, 2), sec_cnt, size=10)
        set_cell_text(t1.cell(r_idx, 3), idl, size=10)
        set_cell_text(t1.cell(r_idx, 4), act_p, size=10)

    # Fill Block 3 (Rows 7~9, Integration)
    set_cell_text(t1.cell(7, 0), "統整活動\n與素養評量", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    t1_block3 = [
        ("統整活動一（篇章閱讀回信、審題立意、朗讀技巧）", "4", "10 ％", "10 ％"),
        ("統整活動二（六何法閱讀策略、寫作選材、讀報應用）", "3", "7 ％", "7 ％"),
        ("總複習與期中評量（全冊統整複習、情境素養整合評量）", "3", "6 ％", "6 ％"),
    ]
    for r_idx, (act, sec_cnt, idl, act_p) in enumerate(t1_block3, start=7):
        set_cell_text(t1.cell(r_idx, 1), act, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(t1.cell(r_idx, 2), sec_cnt, size=10)
        set_cell_text(t1.cell(r_idx, 3), idl, size=10)
        set_cell_text(t1.cell(r_idx, 4), act_p, size=10)

    # Fill Row 10 (Total Row)
    total_row = t1.rows[10]
    set_cell_text(total_row.cells[0], "合　　計", bold=True, size=10.5)
    set_cell_text(total_row.cells[2], "45 節", bold=True, size=10)
    set_cell_text(total_row.cells[3], "100 ％", bold=True, size=10)
    set_cell_text(total_row.cells[4], "100 ％", bold=True, size=10)

    # Add page break right before "二、測驗的題型與配分" so Page 1 has Section 1, Page 2 has Sections 2 & 3
    # Find paragraph "二、測驗的題型與配分"
    for p in doc.paragraphs:
        if "二、測驗的題型與配分" in p.text:
            # Insert page break before this paragraph
            pPr = p._p.get_or_add_pPr()
            pPr.append(parse_xml(f'<w:pageBreakBefore {ns}/>'))
            break

    # 3. Table 2: 測驗的題型與配分
    t2 = doc.tables[1]
    t2_headers = ["題型項目", "一、國字注音", "二、改錯字", "三、選擇題", "四、成語填空", "五、句型練習", "六、閱讀理解"]
    t2_scores = ["配　　分", "28 分", "12 分", "20 分", "10 分", "14 分", "16 分"]
    for c_idx in range(7):
        set_cell_text(t2.rows[0].cells[c_idx], t2_headers[c_idx], bold=True, size=10)
        set_cell_text(t2.rows[1].cells[c_idx], t2_scores[c_idx], bold=True, size=10)

    # 4. Table 3: 試題分析雙向細目表
    t3 = doc.tables[2]
    # Delete unused empty rows in t3 (rows 3 and 4)
    t3._tbl.remove(t3.rows[4]._tr)
    t3._tbl.remove(t3.rows[3]._tr)

    t3_data = [
        ("第一單元：好山好水好故鄉\n（第1~3課、統整一）", "18 分\n12(1)+2(2)+1(2)", "12 分\n2(2)+2(4)", "12 分\n3(2)+1(2)+1(4)", "2 分\n1(2)", "0 分", "3 分\n1(3)", "47 分"),
        ("第二單元：天空的奇想\n（第4~6課、統整二）", "26 分\n16(1)+4(2)+1(2)", "6 分\n3(2)", "8 分\n2(2)+1(4)", "4 分\n1(4)", "6 分\n1(2)+1(4)", "3 分\n1(3)", "53 分"),
    ]
    for r_idx, (u_title, m_mem, m_und, m_app, m_ana, m_eva, m_cre, m_tot) in enumerate(t3_data, start=1):
        row = t3.rows[r_idx]
        set_cell_text(row.cells[0], u_title, bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(row.cells[1], m_mem, size=9)
        set_cell_text(row.cells[2], m_und, size=9)
        set_cell_text(row.cells[3], m_app, size=9)
        set_cell_text(row.cells[4], m_ana, size=9)
        set_cell_text(row.cells[5], m_eva, size=9)
        set_cell_text(row.cells[6], m_cre, size=9)
        set_cell_text(row.cells[7], m_tot, bold=True, size=10)

    # Total row (row 3)
    total_t3 = t3.rows[3]
    set_cell_text(total_t3.cells[0], "合　　計", bold=True, size=10)
    set_cell_text(total_t3.cells[1], "44 分\n(44%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[2], "18 分\n(18%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[3], "20 分\n(20%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[4], "6 分\n(6%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[5], "6 分\n(6%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[6], "6 分\n(6%)", bold=True, size=9.5)
    set_cell_text(total_t3.cells[7], "100 分\n(100%)", bold=True, size=10.5)

    out_docx = r"d:\opencode\115-\A7B5B22356490504BD93DBF276AC179FCE3710A7_附件三：定期評量雙向細目表.docx"
    doc.save(out_docx)
    print("Saved perfect spec docx.")

generate_perfect_spec()
