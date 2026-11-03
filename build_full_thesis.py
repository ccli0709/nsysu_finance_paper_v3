# -*- coding: utf-8 -*-
"""
build_full_thesis.py
--------------------
國立中山大學財務管理學系 碩士論文 完整生成程式
題目：企業水資源管理與廢棄物處理對成本僵固性之影響
The Impact of Corporate Water Resource Management and Waste Treatment on Cost Stickiness

特色：
1. 完整落實《碩士論文要求格式.pdf》與《碩士論文格式參考.docx》規範：
   - 雙軌頁碼系統（前置頁小寫羅馬數字 i, ii, iii...；正文起阿拉伯數字 1, 2, 3...）。
   - 標準版面：上 2.54 cm、下 2.54 cm、左 3.0 cm、右 3.0 cm。
   - 字體：中文 標楷體、英文 Times New Roman，內文 1.5 倍行高、首行縮排 2 字元。
   - 標準三線表（Three-line Table）：頂線 1.5pt、表頭下線 0.75pt、底線 1.5pt，無縱線。
2. 完整落實《論文發展提示詞.md》之學術紀律：
   - 100% 採用台灣財務管理學術用語（成本僵固性、交乘項、叢集穩健標準誤、縮尾處理）。
   - 100% 鎖定真實實證數據（高耗水當期用水密集度 β3 = -0.0914**；高污染遞延一期廢棄物密集度 β3 = -0.0503*）。
   - 100% 限定 20 篇指定參考文獻（中文依筆畫順序、英文依字母順序，APA 7th 體例與 0.85cm 凸排）。
3. 輸出至「論文生成結果/」資料夾，檔名後綴附上生成日期時間。
"""

import os
import sys
from datetime import datetime
import numpy as np
import pandas as pd
import docx
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# 設定資料庫與路徑
BASE_DIR = r"d:\Repositories\nsysu_finance_paper_v3"
DATA_DIR = os.path.join(BASE_DIR, "資料分析結果")
MODEL_CSV = os.path.join(DATA_DIR, "05-模型變數與交乘項建構", "04_model_data.csv")
COEF_CSV = os.path.join(DATA_DIR, "06-實證迴歸模型分析", "05_regression_coef.csv")
ROB_CSV = os.path.join(DATA_DIR, "07-穩健性檢定與分析", "06_robustness_summary.csv")
IND_CSV = os.path.join(DATA_DIR, "04-樣本篩選與產業統計", "03_industry_stats.csv")
WATERUSE_CSV = os.path.join(DATA_DIR, "05-模型變數與交乘項建構", "04_industry_wateruse.csv")
OUT_DIR = os.path.join(BASE_DIR, "論文生成結果")

os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "標楷體"
EN_FONT = "Times New Roman"

def set_run_font(run, font_name_cjk=CJK_FONT, font_name_en=EN_FONT, size_pt=12, bold=False, italic=False, color_rgb=(0,0,0)):
    run.font.name = font_name_en
    rPr = run._element.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name_en}" w:hAnsi="{font_name_en}" w:eastAsia="{font_name_cjk}" w:cs="{font_name_en}"/>')
    rPr.append(rFonts)
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

def set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=0, space_after=0, first_line_indent=0):
    p.alignment = align
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if first_line_indent > 0:
        p.paragraph_format.first_line_indent = Pt(first_line_indent)

def add_styled_para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, bold=False, italic=False, first_indent=False, space_before=0, space_after=0, line_spacing=1.5):
    p = doc.add_paragraph()
    indent_val = 24 if first_indent else 0
    set_para_format(p, align=align, line_spacing=line_spacing, space_before=space_before, space_after=space_after, first_line_indent=indent_val)
    if text:
        run = p.add_run(text)
        set_run_font(run, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_chapter_heading(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before=18, space_after=12)
    run = p.add_run(text)
    set_run_font(run, size_pt=18, bold=True)
    return p

def add_section_heading(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=12, space_after=6)
    run = p.add_run(text)
    set_run_font(run, size_pt=14, bold=True)
    return p

def add_subsection_heading(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=8, space_after=4)
    run = p.add_run(text)
    set_run_font(run, size_pt=12, bold=True)
    return p

def add_body_para(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=0, space_after=0, first_line_indent=24)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        set_run_font(r_pre, size_pt=12, bold=True)
    run = p.add_run(text)
    set_run_font(run, size_pt=12, bold=False)
    return p

def add_callout_para(doc, text, bold_title=None):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=4, space_after=4, first_line_indent=0)
    if bold_title:
        r_t = p.add_run(bold_title)
        set_run_font(r_t, size_pt=12, bold=True)
    run = p.add_run(text)
    set_run_font(run, size_pt=12, bold=True)
    return p

def apply_three_line_table(table, header_rows=1):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)
    
    for r_idx in range(header_rows):
        for cell in table.rows[r_idx].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

def style_cell(cell, text="", size_pt=10, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, vertical_align=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = vertical_align
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)
    
    p = cell.paragraphs[0]
    p.text = ""
    set_para_format(p, align=align, line_spacing=1.1, space_before=0, space_after=0, first_line_indent=0)
    if text:
        run = p.add_run(str(text))
        set_run_font(run, size_pt=size_pt, bold=bold)

def add_table_title(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=12, space_after=4)
    run = p.add_run(text)
    set_run_font(run, size_pt=11, bold=True)
    return p

def add_table_note(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.2, space_before=4, space_after=12)
    run = p.add_run(text)
    set_run_font(run, size_pt=9, bold=False)
    return p

def set_page_number_footer(section, is_roman=False, start_at_1=False):
    footer = section.footer
    p = footer.paragraphs[0]
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    p.text = ""
    fldSimple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
    p._p.append(fldSimple)
    
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    
    if is_roman:
        pgNumType.set(qn('w:fmt'), 'roman')
    else:
        pgNumType.set(qn('w:fmt'), 'decimal')
        if start_at_1:
            pgNumType.set(qn('w:start'), '1')

print("樣式模組載入成功。")

