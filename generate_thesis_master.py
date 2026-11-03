# -*- coding: utf-8 -*-
"""
generate_thesis_master.py
-------------------------
國立中山大學財務管理學系 碩士學位論文 生成主程式
題目：企業水資源管理與廢棄物處理對成本僵固性之影響
The Impact of Corporate Water Resource Management and Waste Treatment on Cost Stickiness

遵循規範：
1. 國立中山大學碩士論文格式規約（碩士論文要求格式.pdf 與 碩士論文格式參考.docx）
2. 論文發展提示詞.md（台灣財管學術用語、零捏造數據、20篇專屬文獻、三線表標準）
3. 輸出至 論文生成結果/ 檔名後綴附上生成日期時間（YYYYMMDD_HHmmss）
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

# 設定路徑
BASE_DIR = r"d:\Repositories\nsysu_finance_paper_v3"
DATA_DIR = os.path.join(BASE_DIR, "資料分析結果")
MODEL_CSV = os.path.join(DATA_DIR, "05-模型變數與交乘項建構", "04_model_data.csv")
COEF_CSV = os.path.join(DATA_DIR, "06-實證迴歸模型分析", "05_regression_coef.csv")
ROB_CSV = os.path.join(DATA_DIR, "07-穩健性檢定與分析", "06_robustness_summary.csv")
IND_CSV = os.path.join(DATA_DIR, "04-樣本篩選與產業統計", "03_industry_stats.csv")
WATERUSE_CSV = os.path.join(DATA_DIR, "05-模型變數與交乘項建構", "04_industry_wateruse.csv")
OUT_DIR = os.path.join(BASE_DIR, "論文生成結果")

# 字體常數
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

def apply_three_line_table(table, header_rows=1):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 設置整個表格的三線表邊框（頂線 1.5pt，底線 1.5pt，無豎線）
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
    
    # 設置表頭下邊線（細線 0.75pt）
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
    # 減少內邊距以利排版
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)
    
    p = cell.paragraphs[0]
    p.text = ""
    set_para_format(p, align=align, line_spacing=1.1, space_before=0, space_after=0, first_line_indent=0)
    if text:
        run = p.add_run(str(text))
        set_run_font(run, size_pt=size_pt, bold=bold)

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

print("基礎樣式設定工具載入完成。")

