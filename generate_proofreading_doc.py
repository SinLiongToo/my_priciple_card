# -*- coding: utf-8 -*-
"""
generate_proofreading_doc.py
========================================================================================
工作的管見（My Two Cents） — 職場生存與成長的避坑指南
三語對照校稿列印手冊生成器 (Word .docx + PDF)

特色：
1. A4 橫向 (Landscape) 表格對照排版，專為紙本列印與筆墨校對設計。
2. 六欄完整對照：編號/原則關聯、投影片原文、優化英文、優化中文、優化台文、校對勾選與批註欄。
3. 表格標題列跨頁自動重複 (tblHeader)，每列內容不跨頁截斷 (cantSplit)。
4. 完整收錄：篇章核心引導金句（53 則）、20 大核心原則、458 條實務原則卡牌。
5. 同步產出 DOCX 與 PDF 雙格式。
========================================================================================
"""

import os
import sys
import json
import subprocess
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

W_NS = nsdecls("w")


def set_cell_background(cell, fill_hex):
    """Set background color for a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("shd"):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {W_NS} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    """Set cell internal padding in twips (20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("tcMar"):
            tcPr.remove(child)
    tcMar = parse_xml(
        f'<w:tcMar {W_NS}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_cell_width(cell, width_in_inches):
    """Set cell explicit width."""
    cell.width = Inches(width_in_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("tcW"):
            tcPr.remove(child)
    dxa_width = int(width_in_inches * 1440)
    tcW = parse_xml(f'<w:tcW {W_NS} w:w="{dxa_width}" w:type="dxa"/>')
    tcPr.append(tcW)


def set_table_borders(table, color="CBD5E1", sz="4"):
    """Apply clean light gray borders to table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {W_NS}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def format_row_protection(row, is_header=False):
    """Set repeat header on header row, and cantSplit on all content rows."""
    trPr = row._tr.get_or_add_trPr()
    if is_header:
        trPr.append(parse_xml(f'<w:tblHeader {W_NS}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {W_NS}/>'))


def set_run_font(run, font_name="Microsoft JhengHei", size_pt=None, bold=None, italic=None, color=None, east_asia_font="Microsoft JhengHei"):
    """Format run font, explicitly setting both Western (ascii/hAnsi) and East Asian (eastAsia) font families.
    This prevents Word from falling back to Western punctuation metrics which visually compresses full-width punctuation.
    """
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), east_asia_font)
    rFonts.set(qn('w:cs'), east_asia_font)

    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def add_page_number_to_section(section):
    """Setup header and footer with dynamic page numbers and metadata."""
    # Header
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun1 = hp.add_run("工作的管見（My Two Cents） — 職場生存與成長的避坑指南【三語對照校稿手冊】")
    set_run_font(hrun1, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(100, 116, 139))

    # Footer
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    frun1 = fp.add_run("□ 未校對    ☑ 校對完成    |    校對日期：______ 年 ___ 月 ___ 日    校對人：__________               ")
    set_run_font(frun1, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(100, 116, 139))

    frun2 = fp.add_run("第 ")
    set_run_font(frun2, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(100, 116, 139))

    fld1 = OxmlElement("w:fldSimple")
    fld1.set(qn("w:instr"), "PAGE")
    fp._p.append(fld1)

    frun3 = fp.add_run(" 頁 / 共 ")
    set_run_font(frun3, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(100, 116, 139))

    fld2 = OxmlElement("w:fldSimple")
    fld2.set(qn("w:instr"), "NUMPAGES")
    fp._p.append(fld2)

    frun4 = fp.add_run(" 頁")
    set_run_font(frun4, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(100, 116, 139))


def get_chap_info(p_num):
    """Return chapter id and title for a given principle number."""
    if p_num <= 4: return "1", "第一篇：效能與工具（生產力管理）"
    if p_num <= 8: return "2", "第二篇：思考與決策（思維模型與問題解決）"
    if p_num <= 11: return "3", "第三篇：溝通與協作（職場對話術）"
    if p_num <= 14: return "4", "第四篇：職場生存與防線（自我保護與邊界）"
    if p_num <= 17: return "5", "第五篇：職涯與個人成長（價值提升與升遷）"
    return "6", "第六篇：能量、韌性與生活平衡（心智與身心管理）"


def build_proofreading_docx(principles, final_items, chapters, docx_path):
    """Generate the landscape proofreading Word document."""
    print(f"Creating Word proofreading document: {docx_path}...")
    doc = docx.Document()

    # 1. Page Setup: A4 Landscape, 0.5 in margins
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)   # 29.7 cm
    section.page_height = Inches(8.27)   # 21.0 cm
    section.top_margin = Inches(0.5)     # 1.27 cm
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    add_page_number_to_section(section)

    # Prevent Word from automatically compressing CJK full-width punctuation
    csc = OxmlElement('w:characterSpacingControl')
    csc.set(qn('w:val'), 'doNotCompress')
    doc.settings._element.append(csc)

    # Set document base font with both Western and EastAsian font names
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Microsoft JhengHei"
    normal_style.font.size = Pt(8.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    rPr = normal_style._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Microsoft JhengHei')
    rFonts.set(qn('w:hAnsi'), 'Microsoft JhengHei')
    rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    rFonts.set(qn('w:cs'), 'Microsoft JhengHei')

    # 2. Cover / Guide Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(4)
    title_p.paragraph_format.space_after = Pt(2)
    title_run = title_p.add_run("工作的管見（My Two Cents） — 職場生存與成長的避坑指南")
    set_run_font(title_run, font_name="Microsoft JhengHei", size_pt=18, bold=True, color=RGBColor(30, 41, 59))

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(6)
    sub_run = sub_p.add_run("From the Floor: Practical Work Principles — 英文・中文・台文 三語對照校稿列印手冊（A4 橫向表格版）")
    set_run_font(sub_run, font_name="Microsoft JhengHei", size_pt=11, bold=True, color=RGBColor(79, 70, 229))

    # Guide Box
    guide_table = doc.add_table(rows=1, cols=1)
    guide_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    guide_cell = guide_table.rows[0].cells[0]
    set_cell_background(guide_cell, "F1F5F9")
    set_cell_margins(guide_cell, top=120, bottom=120, left=180, right=180)
    set_cell_width(guide_cell, 10.69)

    gp = guide_cell.paragraphs[0]
    gp.paragraph_format.space_before = Pt(2)
    gp.paragraph_format.space_after = Pt(2)
    g_title = gp.add_run("【紙本校稿指引與標記標準】\n")
    set_run_font(g_title, font_name="Microsoft JhengHei", size_pt=9.5, bold=True, color=RGBColor(15, 23, 42))

    g_body = gp.add_run(
        "1. 收錄範圍：完整收錄六大篇章核心金句（53 則）、20 大核心原則，以及 458 張實務原則卡牌（Slide 1～Slide 9 全集）。\n"
        "2. 三語核對重點：\n"
        "   • 優化英文 (Polished English)：商務語境自然道地，文法與用語專業簡潔，修飾口語與語病。\n"
        "   • 優化中文 (Polished Mandarin)：繁體中文流暢嚴謹，符合職場實務溝通脈絡，文雅精準。\n"
        "   • 優化台文 (Polished Taiwanese)：採教育部推薦標準漢字（如「鬥相共」、「省氣力」、「心頭抓乎定」），蘊含處世智慧與俗諺感。\n"
        "   • 投影片原文 (Slide Original)：保留簡報原始文字，方便對照原意與優化前後之精準度。\n"
        "3. 紙本校稿操作：每條右側均附「□ 通過 / □ 待修」勾選格與空白手寫筆記空間，可直接使用紅筆進行批註修訂。\n"
        "4. 校對簽核欄：\n"
        "   校對總結：[  ] 全書審定通過    [  ] 需部分二校修訂    簽署：_______________    日期：2026 年 ___ 月 ___ 日\n"
        f"5. 產出時間：{datetime.now().strftime('%Y-%m-%d %H:%M')} | 總卡牌數：458 條 | 核心原則：20 條 | 篇章金句：53 則"
    )
    set_run_font(g_body, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(51, 65, 85))

    # Page break after Cover & Guidelines so Part 1 starts on Page 2
    doc.add_page_break()

    # 3. PART 1: 核心原則與篇章金句 (Principles & Chapter Slogans)
    p1_head = doc.add_paragraph()
    p1_head.paragraph_format.space_before = Pt(8)
    p1_head.paragraph_format.space_after = Pt(4)
    p1_head_run = p1_head.add_run("第一部分：六大篇章核心指引金句與 20 大核心原則")
    set_run_font(p1_head_run, font_name="Microsoft JhengHei", size_pt=13, bold=True, color=RGBColor(30, 41, 59))

    # Group principles by chapter
    chap_groups = {}
    for p in principles:
        cid, ctitle = get_chap_info(p['num'])
        if cid not in chap_groups:
            chap_groups[cid] = {"title": ctitle, "items": []}
        chap_groups[cid]["items"].append(p)

    for cid in sorted(chap_groups.keys()):
        cdata = chap_groups[cid]
        chap_p = doc.add_paragraph()
        chap_p.paragraph_format.space_before = Pt(10)
        chap_p.paragraph_format.space_after = Pt(2)
        c_run = chap_p.add_run(f"■ {cdata['title']}")
        set_run_font(c_run, font_name="Microsoft JhengHei", size_pt=11, bold=True, color=RGBColor(67, 56, 202))

        # Quotes for this chapter
        matched_chap = next((c for c in chapters if c['id'] == cid), None) if chapters else None
        if matched_chap and matched_chap.get('quotes'):
            qp = doc.add_paragraph()
            qp.paragraph_format.space_before = Pt(0)
            qp.paragraph_format.space_after = Pt(4)
            q_lead = qp.add_run("【核心金句】 ")
            set_run_font(q_lead, font_name="Microsoft JhengHei", size_pt=8.5, bold=True, color=RGBColor(180, 83, 9))
            q_txt = qp.add_run("   ".join(matched_chap['quotes']))
            set_run_font(q_txt, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(120, 53, 15))

        # Principles Table for this chapter
        # Columns: 原則編號 (1.3 in) | 優化英文 (2.5 in) | 優化中文 (2.4 in) | 優化台文 (2.4 in) | 對照卡牌 (1.3 in) | 校對批註 (0.79 in)
        p_table = doc.add_table(rows=1, cols=6)
        p_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(p_table)

        col_widths_p = [1.3, 2.5, 2.4, 2.4, 1.3, 0.79]
        hdr_titles_p = ["原則編號與名稱", "優化英文 (English)", "優化中文 (Mandarin)", "優化台文 (Taiwanese)", "對照卡牌編號", "校對批註"]
        
        hdr_row = p_table.rows[0]
        format_row_protection(hdr_row, is_header=True)
        for i, (title, w) in enumerate(zip(hdr_titles_p, col_widths_p)):
            cell = hdr_row.cells[i]
            set_cell_background(cell, "312E81") # Dark Indigo
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            set_cell_width(cell, w)
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cp.add_run(title)
            set_run_font(r, font_name="Microsoft JhengHei", size_pt=8.5, bold=True, color=RGBColor(255, 255, 255))

        for row_idx, p_item in enumerate(cdata["items"]):
            row = p_table.add_row()
            format_row_protection(row, is_header=False)
            bg = "FFFFFF" if row_idx % 2 == 0 else "F8FAFC"

            # Col 0: 原則編號與名稱
            c0 = row.cells[0]
            set_cell_background(c0, bg)
            set_cell_margins(c0)
            set_cell_width(c0, col_widths_p[0])
            p0 = c0.paragraphs[0]
            r0_1 = p0.add_run(f"原則 {p_item['num']}\n")
            set_run_font(r0_1, font_name="Microsoft JhengHei", size_pt=9, bold=True, color=RGBColor(67, 56, 202))
            r0_2 = p0.add_run(p_item['title'])
            set_run_font(r0_2, font_name="Microsoft JhengHei", size_pt=8, color=RGBColor(51, 65, 85))

            # Col 1: 優化英文
            c1 = row.cells[1]
            set_cell_background(c1, bg)
            set_cell_margins(c1)
            set_cell_width(c1, col_widths_p[1])
            p1 = c1.paragraphs[0]
            p1.paragraph_format.line_spacing = 1.15
            r1 = p1.add_run(p_item['english'])
            set_run_font(r1, font_name="Calibri", size_pt=8.5, color=RGBColor(30, 41, 59), east_asia_font="Microsoft JhengHei")

            # Col 2: 優化中文
            c2 = row.cells[2]
            set_cell_background(c2, bg)
            set_cell_margins(c2)
            set_cell_width(c2, col_widths_p[2])
            p2 = c2.paragraphs[0]
            p2.paragraph_format.line_spacing = 1.15
            r2 = p2.add_run(p_item['mandarin'])
            set_run_font(r2, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(15, 23, 42))

            # Col 3: 優化台文
            c3 = row.cells[3]
            set_cell_background(c3, bg)
            set_cell_margins(c3)
            set_cell_width(c3, col_widths_p[3])
            p3 = c3.paragraphs[0]
            p3.paragraph_format.line_spacing = 1.15
            r3 = p3.add_run(p_item['taiwanese'])
            set_run_font(r3, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(15, 118, 110))

            # Col 4: 對照卡牌
            c4 = row.cells[4]
            set_cell_background(c4, bg)
            set_cell_margins(c4)
            set_cell_width(c4, col_widths_p[4])
            p4 = c4.paragraphs[0]
            p4.paragraph_format.line_spacing = 1.1
            r4 = p4.add_run(p_item['ref'])
            set_run_font(r4, font_name="Calibri", size_pt=8, color=RGBColor(100, 116, 139), east_asia_font="Microsoft JhengHei")

            # Col 5: 校對批註
            c5 = row.cells[5]
            set_cell_background(c5, bg)
            set_cell_margins(c5)
            set_cell_width(c5, col_widths_p[5])
            p5 = c5.paragraphs[0]
            p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r5 = p5.add_run("□ 通過\n□ 待修\n──────")
            set_run_font(r5, font_name="Microsoft JhengHei", size_pt=7.5, color=RGBColor(148, 163, 184))

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4. PART 2: 458 張實務原則卡牌 (458 Practice Cards)
    doc.add_page_break()

    p2_head = doc.add_paragraph()
    p2_head.paragraph_format.space_before = Pt(8)
    p2_head.paragraph_format.space_after = Pt(2)
    p2_head_run = p2_head.add_run("第二部分：458 張實務原則卡牌（投影片逐頁對照校對表）")
    p2_head_run.font.name = "Microsoft JhengHei"
    p2_head_run.font.size = Pt(13)
    p2_head_run.font.bold = True
    p2_head_run.font.color.rgb = RGBColor(30, 41, 59)

    p2_sub = doc.add_paragraph()
    p2_sub.paragraph_format.space_before = Pt(0)
    p2_sub.paragraph_format.space_after = Pt(6)
    p2_sub_run = p2_sub.add_run("本表按投影片 Slide 1 至 Slide 9 循序排列，包含原始英文、優化英文、優化中文與優化台文，右側附校對筆記欄。")
    p2_sub_run.font.name = "Microsoft JhengHei"
    p2_sub_run.font.size = Pt(8.5)
    p2_sub_run.font.color.rgb = RGBColor(100, 116, 139)

    # Group final_items by slide
    slide_groups = {}
    for item in final_items:
        s = item['slide']
        if s not in slide_groups:
            slide_groups[s] = []
        slide_groups[s].append(item)

    # Table Column Specification for Part 2:
    # 0: 編號/關聯 (1.05 in)
    # 1: 投影片原文 (2.10 in)
    # 2: 優化英文 (2.30 in)
    # 3: 優化中文 (2.25 in)
    # 4: 優化台文 (2.25 in)
    # 5: 校對狀態與批註 (0.74 in)
    # Sum: 1.05 + 2.10 + 2.30 + 2.25 + 2.25 + 0.74 = 10.69 in
    col_widths_c = [1.05, 2.10, 2.30, 2.25, 2.25, 0.74]
    hdr_titles_c = ["編號 / 關聯", "投影片原文 (Original)", "優化英文 (Polished English)", "優化中文 (Polished Mandarin)", "優化台文 (Polished Taiwanese)", "校對批註"]

    for s_num in sorted(slide_groups.keys()):
        s_items = slide_groups[s_num]
        min_key = s_items[0]['key']
        max_key = s_items[-1]['key']

        # Slide Section Header
        s_header_p = doc.add_paragraph()
        s_header_p.paragraph_format.space_before = Pt(8)
        s_header_p.paragraph_format.space_after = Pt(3)
        sh_run = s_header_p.add_run(f"【投影片 Slide {s_num}】 卡牌 No. {min_key} ～ No. {max_key}（共 {len(s_items)} 條）")
        set_run_font(sh_run, font_name="Microsoft JhengHei", size_pt=10.5, bold=True, color=RGBColor(15, 118, 110))

        # Slide Table
        c_table = doc.add_table(rows=1, cols=6)
        c_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(c_table)

        # Header Row
        hdr_row = c_table.rows[0]
        format_row_protection(hdr_row, is_header=True)
        for i, (title, w) in enumerate(zip(hdr_titles_c, col_widths_c)):
            cell = hdr_row.cells[i]
            set_cell_background(cell, "1E293B") # Dark Slate
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            set_cell_width(cell, w)
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cp.add_run(title)
            set_run_font(r, font_name="Microsoft JhengHei", size_pt=8.5, bold=True, color=RGBColor(255, 255, 255))

        for row_idx, item in enumerate(s_items):
            row = c_table.add_row()
            format_row_protection(row, is_header=False)
            bg = "FFFFFF" if row_idx % 2 == 0 else "F8FAFC"

            # Determine matched principles
            matched_principles = []
            item_key = str(item['key'])
            base_key = item_key.split('-')[0]
            for p in principles:
                refs = [r.strip() for r in p['ref'].split(',')]
                if item_key in refs or base_key in refs:
                    matched_principles.append(p['num'])

            # Col 0: 編號 / 關聯
            c0 = row.cells[0]
            set_cell_background(c0, bg)
            set_cell_margins(c0)
            set_cell_width(c0, col_widths_c[0])
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            r0_num = p0.add_run(f"No. {item['key']}\n")
            set_run_font(r0_num, font_name="Calibri", size_pt=9.5, bold=True, color=RGBColor(30, 41, 59), east_asia_font="Microsoft JhengHei")

            r0_slide = p0.add_run(f"Slide {item['slide']}\n")
            set_run_font(r0_slide, font_name="Calibri", size_pt=7.5, color=RGBColor(148, 163, 184), east_asia_font="Microsoft JhengHei")

            if matched_principles:
                pr_str = ", ".join(str(n) for n in matched_principles)
                r0_pr = p0.add_run(f"原則 {pr_str}")
                set_run_font(r0_pr, font_name="Microsoft JhengHei", size_pt=7.5, color=RGBColor(79, 70, 229))

            # Col 1: 投影片原文
            c1 = row.cells[1]
            set_cell_background(c1, bg)
            set_cell_margins(c1)
            set_cell_width(c1, col_widths_c[1])
            p1 = c1.paragraphs[0]
            p1.paragraph_format.line_spacing = 1.1
            r1 = p1.add_run(item.get('original', ''))
            set_run_font(r1, font_name="Calibri", size_pt=8, italic=True, color=RGBColor(71, 85, 105), east_asia_font="Microsoft JhengHei")

            # Col 2: 優化英文
            c2 = row.cells[2]
            set_cell_background(c2, bg)
            set_cell_margins(c2)
            set_cell_width(c2, col_widths_c[2])
            p2 = c2.paragraphs[0]
            p2.paragraph_format.line_spacing = 1.15
            r2 = p2.add_run(item.get('english', ''))
            set_run_font(r2, font_name="Calibri", size_pt=8.5, color=RGBColor(30, 41, 59), east_asia_font="Microsoft JhengHei")

            # Col 3: 優化中文
            c3 = row.cells[3]
            set_cell_background(c3, bg)
            set_cell_margins(c3)
            set_cell_width(c3, col_widths_c[3])
            p3 = c3.paragraphs[0]
            p3.paragraph_format.line_spacing = 1.15
            r3 = p3.add_run(item.get('mandarin', ''))
            set_run_font(r3, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(15, 23, 42))

            # Col 4: 優化台文
            c4 = row.cells[4]
            set_cell_background(c4, bg)
            set_cell_margins(c4)
            set_cell_width(c4, col_widths_c[4])
            p4 = c4.paragraphs[0]
            p4.paragraph_format.line_spacing = 1.15
            r4 = p4.add_run(item.get('taiwanese', ''))
            set_run_font(r4, font_name="Microsoft JhengHei", size_pt=8.5, color=RGBColor(15, 118, 110))

            # Col 5: 校對狀態與批註
            c5 = row.cells[5]
            set_cell_background(c5, bg)
            set_cell_margins(c5)
            set_cell_width(c5, col_widths_c[5])
            p5 = c5.paragraphs[0]
            p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r5 = p5.add_run("□ 通過\n□ 待修\n──────")
            set_run_font(r5, font_name="Microsoft JhengHei", size_pt=7.5, color=RGBColor(148, 163, 184))

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Save DOCX
    doc.save(docx_path)
    print(f"Successfully generated DOCX: {docx_path}")


def convert_docx_to_pdf(docx_path, output_dir="."):
    """Convert generated DOCX to PDF using headless LibreOffice."""
    soffice_path = r"C:\Program Files\LibreOffice\program\soffice.com"
    if not os.path.exists(soffice_path):
        # Fallback to soffice in PATH
        soffice_path = "soffice.com"

    abs_docx = os.path.abspath(docx_path)
    abs_outdir = os.path.abspath(output_dir)
    print(f"Converting DOCX to PDF using LibreOffice: {abs_docx} -> {abs_outdir}...")

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        abs_docx,
        "--outdir", abs_outdir
    ]

    try:
        res = subprocess.run(cmd, capture_output=True, errors="replace", timeout=120)
        pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        expected_pdf = os.path.join(output_dir, pdf_name)
        if os.path.exists(expected_pdf) and os.path.getsize(expected_pdf) > 0:
            print(f"Successfully generated PDF: {expected_pdf} ({os.path.getsize(expected_pdf):,} bytes)")
            return expected_pdf
        else:
            print(f"Warning: PDF conversion completed with code {res.returncode}, but PDF not found or empty.")
            if res.stdout: print("stdout:", res.stdout)
            if res.stderr: print("stderr:", res.stderr)
            return None
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
        return None


def build_proofreading_xlsx(principles, final_items, chapters, xlsx_path):
    """Generate high-aesthetic, formula-driven Excel proofreading workbook (.xlsx)."""
    print(f"Creating Excel proofreading workbook: {xlsx_path}...")
    wb = openpyxl.Workbook()

    # Common styles
    font_family = "Microsoft JhengHei"
    font_en_family = "Calibri"

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    thick_bottom_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='medium', color='1E293B')
    )

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Color fills
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # Status styles for conditional formatting
    fill_pass = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_pass = Font(name=font_family, size=9.5, bold=True, color="166534")

    fill_fix = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_fix = Font(name=font_family, size=9.5, bold=True, color="991B1B")

    fill_revised = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
    font_revised = Font(name=font_family, size=9.5, bold=True, color="3730A3")

    fill_question = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_question = Font(name=font_family, size=9.5, bold=True, color="92400E")

    def apply_status_conditional_formatting(ws, cell_range):
        ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"通過"'], stopIfTrue=True, fill=fill_pass, font=font_pass))
        ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"待修"'], stopIfTrue=True, fill=fill_fix, font=font_fix))
        ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"已修訂"'], stopIfTrue=True, fill=fill_revised, font=font_revised))
        ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"有疑問"'], stopIfTrue=True, fill=fill_question, font=font_question))

    # =========================================================================
    # SHEET 1: 校稿總覽與統計 (Dashboard)
    # =========================================================================
    ws_dash = wb.active
    ws_dash.title = "校稿總覽與統計"
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("B2:I2")
    cell_b2 = ws_dash["B2"]
    cell_b2.value = "工作的管見（My Two Cents） — 職場生存與成長的避坑指南"
    cell_b2.font = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    cell_b2.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    cell_b2.alignment = Alignment(horizontal='center', vertical='center')
    ws_dash.row_dimensions[2].height = 36

    ws_dash.merge_cells("B3:I3")
    cell_b3 = ws_dash["B3"]
    cell_b3.value = f"英文・中文・台文 三語對照校稿手冊（Excel 試算校對版） | 產出時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}"
    cell_b3.font = Font(name=font_family, size=10, color="CBD5E1")
    cell_b3.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    cell_b3.alignment = Alignment(horizontal='center', vertical='center')
    ws_dash.row_dimensions[3].height = 24

    # Instructions box
    ws_dash.merge_cells("B5:I5")
    cell_b5 = ws_dash["B5"]
    cell_b5.value = "【校稿操作指引與使用說明】"
    cell_b5.font = Font(name=font_family, size=11, bold=True, color="0F172A")
    cell_b5.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    cell_b5.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws_dash.row_dimensions[5].height = 25

    instructions = [
        ("B6:I6", "1. 本校稿工作表完整收錄全書 458 條實務原則卡牌、20 大核心原則與 53 則篇章金句，專為三語精準對照、審定與批註修訂設計。"),
        ("B7:I7", "2. 校稿狀態選單：請在各工作表之「校對狀態」欄位點選下拉選單，選取：【通過】、【待修】、【已修訂】、【有疑問】、【未校對】。"),
        ("B8:I8", "3. 狀態顏色高亮：選取狀態後儲存格自動變色（綠色=通過、紅色=待修、藍色=已修訂、黃色=有疑問），視覺化追蹤極為直覺。"),
        ("B9:I9", "4. 即時統計儀表板：下方表格會依據您在各工作表填寫的校對狀態，自動透過 Excel 公式計算通過率與剩餘待審項目。"),
        ("B10:I10", "5. 凍結窗格與篩選：各工作表首列已鎖定標題列，並已啟用自動篩選 (AutoFilter)，方便隨時依投影片或校對狀態快速篩選。")
    ]
    for cell_range, text in instructions:
        ws_dash.merge_cells(cell_range)
        top_cell = ws_dash[cell_range.split(":")[0]]
        top_cell.value = text
        top_cell.font = Font(name=font_family, size=9.5, color="334155")
        top_cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        top_cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws_dash.row_dimensions[int(cell_range[1:cell_range.index(':')])].height = 22

    # Border for instructions
    for r in range(5, 11):
        for c in range(2, 10):
            cell = ws_dash.cell(row=r, column=c)
            b_top = Side(style='thin', color='CBD5E1') if r == 5 else None
            b_bottom = Side(style='thin', color='CBD5E1') if r == 10 else None
            b_left = Side(style='thin', color='CBD5E1') if c == 2 else None
            b_right = Side(style='thin', color='CBD5E1') if c == 9 else None
            cell.border = Border(top=b_top, bottom=b_bottom, left=b_left, right=b_right)

    # Dashboard Summary Table
    ws_dash.merge_cells("B12:I12")
    cell_b12 = ws_dash["B12"]
    cell_b12.value = "■ 全書三語校對進度儀表板 (即時公式動態統計)"
    cell_b12.font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    cell_b12.fill = PatternFill(start_color="312E81", end_color="312E81", fill_type="solid")
    cell_b12.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws_dash.row_dimensions[12].height = 28

    headers_dash = [
        ("項目類別 (Category)", 28),
        ("收錄總數 (Total)", 16),
        ("審定通過 (Passed)", 16),
        ("待修訂 (Needs Fix)", 16),
        ("已修訂 (Revised)", 16),
        ("有疑問 (Questions)", 16),
        ("未校對 (Unchecked)", 16),
        ("完成進度 (Progress)", 18),
    ]

    ws_dash.row_dimensions[13].height = 28
    for col_idx, (h_title, w) in enumerate(headers_dash, start=2):
        cell = ws_dash.cell(row=13, column=col_idx, value=h_title)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4338CA", end_color="4338CA", fill_type="solid")
        cell.alignment = align_center
        cell.border = thin_border

    # Rows in Dashboard Table
    dash_rows = [
        (14, "實務原則卡牌（458 條）",
         "=COUNTA('實務原則卡牌校對 (458條)'!B2:B459)",
         '=COUNTIF(\'實務原則卡牌校對 (458條)\'!H2:H459, "通過")',
         '=COUNTIF(\'實務原則卡牌校對 (458條)\'!H2:H459, "待修")',
         '=COUNTIF(\'實務原則卡牌校對 (458條)\'!H2:H459, "已修訂")',
         '=COUNTIF(\'實務原則卡牌校對 (458條)\'!H2:H459, "有疑問")',
         '=C14-D14-E14-F14-G14',
         '=(D14+F14)/C14'),
        (15, "20 大核心原則",
         "=COUNTA('20大核心原則校對'!C2:C21)",
         '=COUNTIF(\'20大核心原則校對\'!I2:I21, "通過")',
         '=COUNTIF(\'20大核心原則校對\'!I2:I21, "待修")',
         '=COUNTIF(\'20大核心原則校對\'!I2:I21, "已修訂")',
         '=COUNTIF(\'20大核心原則校對\'!I2:I21, "有疑問")',
         '=C15-D15-E15-F15-G15',
         '=(D15+F15)/C15'),
        (16, "篇章核心指引金句（53 則）",
         "=COUNTA('篇章核心指引金句 (53則)'!C2:C54)",
         '=COUNTIF(\'篇章核心指引金句 (53則)\'!E2:E54, "通過")',
         '=COUNTIF(\'篇章核心指引金句 (53則)\'!E2:E54, "待修")',
         '=COUNTIF(\'篇章核心指引金句 (53則)\'!E2:E54, "已修訂")',
         '=COUNTIF(\'篇章核心指引金句 (53則)\'!E2:E54, "有疑問")',
         '=C16-D16-E16-F16-G16',
         '=(D16+F16)/C16'),
        (17, "全書合計 (Overall Total)",
         "=SUM(C14:C16)",
         "=SUM(D14:D16)",
         "=SUM(E14:E16)",
         "=SUM(F14:F16)",
         "=SUM(G14:G16)",
         "=SUM(H14:H16)",
         "=(D17+F17)/C17")
    ]

    for row_num, cat_title, f_tot, f_pass, f_fix, f_rev, f_q, f_unchk, f_prog in dash_rows:
        ws_dash.row_dimensions[row_num].height = 25
        is_total = (row_num == 17)
        bg = PatternFill(start_color="EEF2F6" if is_total else ("FFFFFF" if row_num % 2 == 0 else "F8FAFC"),
                         end_color="EEF2F6" if is_total else ("FFFFFF" if row_num % 2 == 0 else "F8FAFC"),
                         fill_type="solid")
        font_c = Font(name=font_family, size=9.5, bold=is_total, color="0F172A" if not is_total else "1E293B")

        # Category name
        c_cat = ws_dash.cell(row=row_num, column=2, value=cat_title)
        c_cat.font = font_c
        c_cat.fill = bg
        c_cat.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c_cat.border = thick_bottom_border if is_total else thin_border

        # Formula cells
        formulas = [f_tot, f_pass, f_fix, f_rev, f_q, f_unchk, f_prog]
        for col_i, f_val in enumerate(formulas, start=3):
            cell = ws_dash.cell(row=row_num, column=col_i, value=f_val)
            cell.font = font_c
            cell.fill = bg
            cell.border = thick_bottom_border if is_total else thin_border
            if col_i == 9: # Progress %
                cell.number_format = '0.0%'
                cell.alignment = align_center
                if is_total:
                    cell.font = Font(name=font_family, size=10, bold=True, color="4338CA")
            else:
                cell.alignment = align_center
                cell.number_format = '#,##0'

    # Set column widths for Dashboard
    ws_dash.column_dimensions['A'].width = 3
    for col_idx, (_, w) in enumerate(headers_dash, start=2):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_dash.column_dimensions[col_letter].width = w

    # Quick Jump Links Box
    ws_dash.merge_cells("B19:I19")
    cell_b19 = ws_dash["B19"]
    cell_b19.value = "【工作表快速導覽】"
    cell_b19.font = Font(name=font_family, size=10, bold=True, color="1E293B")
    cell_b19.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    nav_hints = [
        ("B20:I20", "• 『實務原則卡牌校對 (458條)』工作表：完整 458 條原則卡牌，包含投影片頁次、原文、優化英中台三語與批註欄。"),
        ("B21:I21", "• 『20大核心原則校對』工作表：全書 20 大核心原則，依六大篇章分類對照，含對應實務卡牌編號清單。"),
        ("B22:I22", "• 『篇章核心指引金句 (53則)』工作表：六大篇章之 53 則引導星號金句，審閱精準度與文化語韻。")
    ]
    for cell_range, text in nav_hints:
        ws_dash.merge_cells(cell_range)
        top_cell = ws_dash[cell_range.split(":")[0]]
        top_cell.value = text
        top_cell.font = Font(name=font_family, size=9, color="64748B")
        top_cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws_dash.row_dimensions[int(cell_range[1:cell_range.index(':')])].height = 20

    # =========================================================================
    # SHEET 2: 實務原則卡牌校對 (458條)
    # =========================================================================
    ws_cards = wb.create_sheet(title="實務原則卡牌校對 (458條)")
    ws_cards.views.sheetView[0].showGridLines = True
    ws_cards.freeze_panes = "A2"

    headers_cards = [
        ("投影片", 11),
        ("卡牌編號", 14),
        ("關聯核心原則", 18),
        ("投影片原文 (Slide Original)", 44),
        ("優化英文 (Polished English)", 48),
        ("優化中文 (Polished Mandarin)", 44),
        ("優化台文 (Polished Taiwanese)", 44),
        ("校對狀態", 14),
        ("修訂建議與批註", 36),
        ("校對人", 13),
        ("校對日期", 14)
    ]

    ws_cards.row_dimensions[1].height = 30
    for col_idx, (h_title, w) in enumerate(headers_cards, start=1):
        cell = ws_cards.cell(row=1, column=col_idx, value=h_title)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_cards.column_dimensions[col_letter].width = w

    # Populate cards
    dv_status_cards = DataValidation(type="list", formula1='"通過,待修,已修訂,有疑問,未校對"', allow_blank=True)
    dv_status_cards.error = "請選擇有效的校對狀態：通過, 待修, 已修訂, 有疑問, 未校對"
    dv_status_cards.errorTitle = "輸入無效"
    dv_status_cards.prompt = "請選取校對狀態"
    dv_status_cards.promptTitle = "狀態選單"
    ws_cards.add_data_validation(dv_status_cards)

    for row_idx, item in enumerate(final_items, start=2):
        ws_cards.row_dimensions[row_idx].height = 24
        bg = fill_white if (row_idx % 2 == 0) else fill_zebra

        # Matched principles
        matched_principles = []
        item_key = str(item['key'])
        base_key = item_key.split('-')[0]
        for p in principles:
            refs = [r.strip() for r in p['ref'].split(',')]
            if item_key in refs or base_key in refs:
                matched_principles.append(p['num'])
        ref_str = ", ".join(f"原則 {n}" for n in matched_principles) if matched_principles else ""

        # Col A: Slide
        c_slide = ws_cards.cell(row=row_idx, column=1, value=f"Slide {item['slide']}")
        c_slide.font = Font(name=font_en_family, size=9.5, color="64748B")
        c_slide.alignment = align_center
        c_slide.fill = bg
        c_slide.border = thin_border

        # Col B: Card No
        c_key = ws_cards.cell(row=row_idx, column=2, value=f"No. {item['key']}")
        c_key.font = Font(name=font_en_family, size=9.5, bold=True, color="0F172A")
        c_key.alignment = align_center
        c_key.fill = bg
        c_key.border = thin_border

        # Col C: Principle Ref
        c_ref = ws_cards.cell(row=row_idx, column=3, value=ref_str)
        c_ref.font = Font(name=font_family, size=9, bold=bool(ref_str), color="4338CA" if ref_str else "94A3B8")
        c_ref.alignment = align_center
        c_ref.fill = bg
        c_ref.border = thin_border

        # Col D: Original
        c_orig = ws_cards.cell(row=row_idx, column=4, value=item['original'])
        c_orig.font = Font(name=font_family, size=9, color="334155")
        c_orig.alignment = align_left
        c_orig.fill = bg
        c_orig.border = thin_border

        # Col E: English
        c_eng = ws_cards.cell(row=row_idx, column=5, value=item['english'])
        c_eng.font = Font(name=font_en_family, size=9.5, color="1E293B")
        c_eng.alignment = align_left
        c_eng.fill = bg
        c_eng.border = thin_border

        # Col F: Mandarin
        c_man = ws_cards.cell(row=row_idx, column=6, value=item['mandarin'])
        c_man.font = Font(name=font_family, size=9.5, color="0F172A")
        c_man.alignment = align_left
        c_man.fill = bg
        c_man.border = thin_border

        # Col G: Taiwanese
        c_tai = ws_cards.cell(row=row_idx, column=7, value=item['taiwanese'])
        c_tai.font = Font(name=font_family, size=9.5, color="0F766E") # Deep Teal
        c_tai.alignment = align_left
        c_tai.fill = bg
        c_tai.border = thin_border

        # Col H: Status (Initial blank)
        c_status = ws_cards.cell(row=row_idx, column=8, value="")
        c_status.font = Font(name=font_family, size=9.5, color="64748B")
        c_status.alignment = align_center
        c_status.fill = bg
        c_status.border = thin_border
        dv_status_cards.add(c_status)

        # Col I: Notes
        c_note = ws_cards.cell(row=row_idx, column=9, value="")
        c_note.font = Font(name=font_family, size=9, color="334155")
        c_note.alignment = align_left
        c_note.fill = bg
        c_note.border = thin_border

        # Col J: Reviewer
        c_rev = ws_cards.cell(row=row_idx, column=10, value="")
        c_rev.font = Font(name=font_family, size=9, color="334155")
        c_rev.alignment = align_center
        c_rev.fill = bg
        c_rev.border = thin_border

        # Col K: Date
        c_dt = ws_cards.cell(row=row_idx, column=11, value="")
        c_dt.font = Font(name=font_en_family, size=9, color="334155")
        c_dt.alignment = align_center
        c_dt.fill = bg
        c_dt.border = thin_border

    total_card_rows = len(final_items) + 1
    apply_status_conditional_formatting(ws_cards, f"H2:H{total_card_rows}")
    ws_cards.auto_filter.ref = f"A1:K{total_card_rows}"

    # =========================================================================
    # SHEET 3: 20大核心原則校對
    # =========================================================================
    ws_p = wb.create_sheet(title="20大核心原則校對")
    ws_p.views.sheetView[0].showGridLines = True
    ws_p.freeze_panes = "A2"

    headers_p = [
        ("篇章編號", 11),
        ("所屬篇章名稱", 28),
        ("原則編號", 12),
        ("原則中文名稱 (Title)", 32),
        ("優化英文 (Polished English)", 48),
        ("優化中文 (Polished Mandarin)", 48),
        ("優化台文 (Polished Taiwanese)", 48),
        ("對照卡牌編號 (Card Refs)", 30),
        ("校對狀態", 14),
        ("修訂建議與批註", 36),
        ("校對人", 13),
        ("校對日期", 14)
    ]

    ws_p.row_dimensions[1].height = 30
    for col_idx, (h_title, w) in enumerate(headers_p, start=1):
        cell = ws_p.cell(row=1, column=col_idx, value=h_title)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="312E81", end_color="312E81", fill_type="solid")
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_p.column_dimensions[col_letter].width = w

    dv_status_p = DataValidation(type="list", formula1='"通過,待修,已修訂,有疑問,未校對"', allow_blank=True)
    ws_p.add_data_validation(dv_status_p)

    for row_idx, p in enumerate(principles, start=2):
        ws_p.row_dimensions[row_idx].height = 26
        bg = fill_white if (row_idx % 2 == 0) else fill_zebra
        cid, ctitle = get_chap_info(p['num'])

        # Col A: Chapter ID
        c_cid = ws_p.cell(row=row_idx, column=1, value=f"第 {cid} 篇")
        c_cid.font = Font(name=font_family, size=9, color="64748B")
        c_cid.alignment = align_center
        c_cid.fill = bg
        c_cid.border = thin_border

        # Col B: Chapter Title
        c_ctitle = ws_p.cell(row=row_idx, column=2, value=ctitle)
        c_ctitle.font = Font(name=font_family, size=9, color="475569")
        c_ctitle.alignment = align_left
        c_ctitle.fill = bg
        c_ctitle.border = thin_border

        # Col C: Principle No
        c_pnum = ws_p.cell(row=row_idx, column=3, value=f"原則 {p['num']}")
        c_pnum.font = Font(name=font_family, size=9.5, bold=True, color="312E81")
        c_pnum.alignment = align_center
        c_pnum.fill = bg
        c_pnum.border = thin_border

        # Col D: Title
        c_ptitle = ws_p.cell(row=row_idx, column=4, value=p['title'])
        c_ptitle.font = Font(name=font_family, size=9.5, bold=True, color="0F172A")
        c_ptitle.alignment = align_left
        c_ptitle.fill = bg
        c_ptitle.border = thin_border

        # Col E: English
        c_peng = ws_p.cell(row=row_idx, column=5, value=p['english'])
        c_peng.font = Font(name=font_en_family, size=9.5, color="1E293B")
        c_peng.alignment = align_left
        c_peng.fill = bg
        c_peng.border = thin_border

        # Col F: Mandarin
        c_pman = ws_p.cell(row=row_idx, column=6, value=p['mandarin'])
        c_pman.font = Font(name=font_family, size=9.5, color="0F172A")
        c_pman.alignment = align_left
        c_pman.fill = bg
        c_pman.border = thin_border

        # Col G: Taiwanese
        c_ptai = ws_p.cell(row=row_idx, column=7, value=p['taiwanese'])
        c_ptai.font = Font(name=font_family, size=9.5, color="0F766E")
        c_ptai.alignment = align_left
        c_ptai.fill = bg
        c_ptai.border = thin_border

        # Col H: Card Refs
        c_pref = ws_p.cell(row=row_idx, column=8, value=p['ref'])
        c_pref.font = Font(name=font_en_family, size=8.5, color="64748B")
        c_pref.alignment = align_left
        c_pref.fill = bg
        c_pref.border = thin_border

        # Col I: Status
        c_pstat = ws_p.cell(row=row_idx, column=9, value="")
        c_pstat.font = Font(name=font_family, size=9.5, color="64748B")
        c_pstat.alignment = align_center
        c_pstat.fill = bg
        c_pstat.border = thin_border
        dv_status_p.add(c_pstat)

        # Col J: Notes
        c_pnote = ws_p.cell(row=row_idx, column=10, value="")
        c_pnote.font = Font(name=font_family, size=9, color="334155")
        c_pnote.alignment = align_left
        c_pnote.fill = bg
        c_pnote.border = thin_border

        # Col K: Reviewer
        c_prev = ws_p.cell(row=row_idx, column=11, value="")
        c_prev.font = Font(name=font_family, size=9, color="334155")
        c_prev.alignment = align_center
        c_prev.fill = bg
        c_prev.border = thin_border

        # Col L: Date
        c_pdt = ws_p.cell(row=row_idx, column=12, value="")
        c_pdt.font = Font(name=font_en_family, size=9, color="334155")
        c_pdt.alignment = align_center
        c_pdt.fill = bg
        c_pdt.border = thin_border

    total_p_rows = len(principles) + 1
    apply_status_conditional_formatting(ws_p, f"I2:I{total_p_rows}")
    ws_p.auto_filter.ref = f"A1:L{total_p_rows}"

    # =========================================================================
    # SHEET 4: 篇章核心指引金句 (53則)
    # =========================================================================
    ws_q = wb.create_sheet(title="篇章核心指引金句 (53則)")
    ws_q.views.sheetView[0].showGridLines = True
    ws_q.freeze_panes = "A2"

    headers_q = [
        ("篇章編號", 11),
        ("所屬篇章名稱", 28),
        ("金句編號", 12),
        ("核心指引金句 (Chapter Quote)", 45),
        ("校對狀態", 14),
        ("修訂建議與批註", 36),
        ("校對人", 13),
        ("校對日期", 14)
    ]

    ws_q.row_dimensions[1].height = 30
    for col_idx, (h_title, w) in enumerate(headers_q, start=1):
        cell = ws_q.cell(row=1, column=col_idx, value=h_title)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="78350F", end_color="78350F", fill_type="solid")
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_q.column_dimensions[col_letter].width = w

    dv_status_q = DataValidation(type="list", formula1='"通過,待修,已修訂,有疑問,未校對"', allow_blank=True)
    ws_q.add_data_validation(dv_status_q)

    q_row_idx = 2
    for chap in chapters:
        cid = chap['id']
        ctitle = chap['title']
        quotes = chap.get('quotes', [])
        for q_sub_idx, quote in enumerate(quotes, start=1):
            ws_q.row_dimensions[q_row_idx].height = 25
            bg = fill_white if (q_row_idx % 2 == 0) else fill_zebra

            # Col A: Chapter ID
            c_cid = ws_q.cell(row=q_row_idx, column=1, value=f"第 {cid} 篇")
            c_cid.font = Font(name=font_family, size=9, color="64748B")
            c_cid.alignment = align_center
            c_cid.fill = bg
            c_cid.border = thin_border

            # Col B: Chapter Title
            c_ctitle = ws_q.cell(row=q_row_idx, column=2, value=ctitle)
            c_ctitle.font = Font(name=font_family, size=9, color="475569")
            c_ctitle.alignment = align_left
            c_ctitle.fill = bg
            c_ctitle.border = thin_border

            # Col C: Quote No
            c_qnum = ws_q.cell(row=q_row_idx, column=3, value=f"金句 {q_sub_idx}")
            c_qnum.font = Font(name=font_family, size=9.5, bold=True, color="78350F")
            c_qnum.alignment = align_center
            c_qnum.fill = bg
            c_qnum.border = thin_border

            # Col D: Quote
            c_qtxt = ws_q.cell(row=q_row_idx, column=4, value=quote)
            c_qtxt.font = Font(name=font_family, size=10, bold=True, color="92400E")
            c_qtxt.alignment = align_left
            c_qtxt.fill = bg
            c_qtxt.border = thin_border

            # Col E: Status
            c_qstat = ws_q.cell(row=q_row_idx, column=5, value="")
            c_qstat.font = Font(name=font_family, size=9.5, color="64748B")
            c_qstat.alignment = align_center
            c_qstat.fill = bg
            c_qstat.border = thin_border
            dv_status_q.add(c_qstat)

            # Col F: Notes
            c_qnote = ws_q.cell(row=q_row_idx, column=6, value="")
            c_qnote.font = Font(name=font_family, size=9, color="334155")
            c_qnote.alignment = align_left
            c_qnote.fill = bg
            c_qnote.border = thin_border

            # Col G: Reviewer
            c_qrev = ws_q.cell(row=q_row_idx, column=7, value="")
            c_qrev.font = Font(name=font_family, size=9, color="334155")
            c_qrev.alignment = align_center
            c_qrev.fill = bg
            c_qrev.border = thin_border

            # Col H: Date
            c_qdt = ws_q.cell(row=q_row_idx, column=8, value="")
            c_qdt.font = Font(name=font_en_family, size=9, color="334155")
            c_qdt.alignment = align_center
            c_qdt.fill = bg
            c_qdt.border = thin_border

            q_row_idx += 1

    total_q_rows = q_row_idx - 1
    apply_status_conditional_formatting(ws_q, f"E2:E{total_q_rows}")
    ws_q.auto_filter.ref = f"A1:H{total_q_rows}"

    # Save Excel Workbook
    wb.save(xlsx_path)
    print(f"Successfully generated Excel proofreading workbook: {xlsx_path} ({os.path.getsize(xlsx_path):,} bytes)")
    return xlsx_path


def sync_excel_to_overrides(xlsx_path="工作的管見_三語對照校稿手冊.xlsx", overrides_path="manual_overrides.json"):
    """
    Reverse synchronization helper:
    Inspects Excel proofreading workbook, and if items are marked as '已修訂' or have revised text,
    syncs them into manual_overrides.json so they take top priority in subsequent builds.
    """
    if not os.path.exists(xlsx_path):
        print(f"Error: Excel file not found: {xlsx_path}")
        return False

    print(f"Syncing reviewed changes from Excel: {xlsx_path} -> {overrides_path}...")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    overrides = {}
    if os.path.exists(overrides_path):
        with open(overrides_path, "r", encoding="utf-8-sig") as f:
            overrides = json.load(f)

    updated_count = 0

    # 1. Check Practice Cards
    if "實務原則卡牌校對 (458條)" in wb.sheetnames:
        ws_cards = wb["實務原則卡牌校對 (458條)"]
        for row in range(2, ws_cards.max_row + 1):
            key_val = str(ws_cards.cell(row=row, column=2).value or "").strip()
            if not key_val:
                continue
            clean_key = key_val.replace("No.", "").strip()

            eng = str(ws_cards.cell(row=row, column=5).value or "").strip()
            man = str(ws_cards.cell(row=row, column=6).value or "").strip()
            tai = str(ws_cards.cell(row=row, column=7).value or "").strip()
            status = str(ws_cards.cell(row=row, column=8).value or "").strip()
            note = str(ws_cards.cell(row=row, column=9).value or "").strip()

            if status in ["已修訂", "通過"] and (clean_key in overrides or note or status == "已修訂"):
                if status == "已修訂":
                    if clean_key not in overrides:
                        overrides[clean_key] = {}
                    if eng: overrides[clean_key]["english"] = eng
                    if man: overrides[clean_key]["mandarin"] = man
                    if tai: overrides[clean_key]["taiwanese"] = tai
                    updated_count += 1

    # 2. Check Principles
    if "20大核心原則校對" in wb.sheetnames:
        ws_p = wb["20大核心原則校對"]
        for row in range(2, ws_p.max_row + 1):
            pnum_val = str(ws_p.cell(row=row, column=3).value or "").strip()
            if not pnum_val:
                continue
            clean_pnum = pnum_val.replace("原則", "").strip()
            p_key = f"principle_{clean_pnum}"

            eng = str(ws_p.cell(row=row, column=5).value or "").strip()
            man = str(ws_p.cell(row=row, column=6).value or "").strip()
            tai = str(ws_p.cell(row=row, column=7).value or "").strip()
            status = str(ws_p.cell(row=row, column=9).value or "").strip()

            if status == "已修訂":
                if p_key not in overrides:
                    overrides[p_key] = {}
                if eng: overrides[p_key]["english"] = eng
                if man: overrides[p_key]["mandarin"] = man
                if tai: overrides[p_key]["taiwanese"] = tai
                updated_count += 1

    with open(overrides_path, "w", encoding="utf-8") as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)

    print(f"Sync completed! Updated {updated_count} items into {overrides_path}.")
    return True


def generate_proofreading_documents(principles, final_items, chapters, output_dir="."):
    """Generate Word (.docx), PDF (.pdf), and Excel (.xlsx) proofreading handbooks."""
    docx_path = os.path.join(output_dir, "工作的管見_三語對照校稿手冊.docx")
    build_proofreading_docx(principles, final_items, chapters, docx_path)
    pdf_path = convert_docx_to_pdf(docx_path, output_dir=output_dir)

    xlsx_path = os.path.join(output_dir, "工作的管見_三語對照校稿手冊.xlsx")
    build_proofreading_xlsx(principles, final_items, chapters, xlsx_path)

    return docx_path, pdf_path, xlsx_path


def main():
    """Standalone entry point."""
    if "--import-excel" in sys.argv or "-i" in sys.argv:
        sync_excel_to_overrides()
        return

    import update_handbook

    pptx_path = "工作原則-書.pptx.pptx"
    md_path = "工作原則整理與潤稿.md"
    db_path = "translation_db.json"
    overrides_path = "manual_overrides.json"

    print("Step 1: Parsing PPTX items...")
    pptx_items = update_handbook.parse_pptx_items(pptx_path)

    print("Step 2: Parsing Principles and Chapters from Markdown...")
    _, principles, chapters = update_handbook.parse_existing_md(md_path)

    print("Step 3: Loading translations & manual overrides...")
    db = {}
    if os.path.exists(db_path):
        with open(db_path, "r", encoding="utf-8-sig") as f:
            db = json.load(f)

    overrides = {}
    if os.path.exists(overrides_path):
        try:
            with open(overrides_path, "r", encoding="utf-8-sig") as f:
                raw_overrides = json.load(f)
                overrides = {k: v for k, v in raw_overrides.items() if not k.startswith("_")}
        except Exception as e:
            print(f"Warning: {e}")

    final_items = []
    for item in pptx_items:
        key = item['key']
        orig = item['text']

        override_match = overrides.get(str(key)) or overrides.get(orig)
        db_match = db.get(orig)

        base_eng = db_match['english'] if db_match else orig
        base_man = db_match['mandarin'] if db_match else ""
        base_tai = db_match['taiwanese'] if db_match else ""

        if override_match and isinstance(override_match, dict):
            eng = override_match.get('english', base_eng)
            man = override_match.get('mandarin', base_man)
            tai = override_match.get('taiwanese', base_tai)
        else:
            eng, man, tai = base_eng, base_man, base_tai

        final_items.append({
            "key": key,
            "slide": item['slide'],
            "original": orig,
            "english": eng,
            "mandarin": man,
            "taiwanese": tai
        })

    print("Step 4: Generating Proofreading Handbooks (DOCX, PDF & XLSX)...")
    generate_proofreading_documents(principles, final_items, chapters)
    print("All proofreading documents generated successfully!")


if __name__ == "__main__":
    main()

