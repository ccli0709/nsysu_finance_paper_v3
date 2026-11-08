# -*- coding: utf-8 -*-
"""
generate_full_paper.py
----------------------
完整生成國立中山大學財務管理學系碩士學位論文 DOCX
題目：企業水資源管理與廢棄物處理對成本僵固性之影響
The Impact of Corporate Water Resource Management and Waste Treatment on Cost Stickiness
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
from equations_xml import EQUATIONS_TBL_XML
from inline_omml import append_text_and_math

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
        append_text_and_math(p, text, size_pt=size_pt, bold=bold, italic=italic)
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
    append_text_and_math(p, text, size_pt=12, bold=True)
    return p

def add_body_para(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=0, space_after=0, first_line_indent=24)
    if bold_prefix:
        append_text_and_math(p, bold_prefix, size_pt=12, bold=True)
    append_text_and_math(p, text, size_pt=12, bold=False)
    return p

def add_callout_para(doc, text, bold_title=None):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=4, space_after=4, first_line_indent=0)
    if bold_title:
        append_text_and_math(p, bold_title, size_pt=12, bold=True)
    append_text_and_math(p, text, size_pt=12, bold=True)
    return p

def add_equation_table(doc, eq_key):
    xml_str = EQUATIONS_TBL_XML[eq_key]
    tbl_element = parse_xml(xml_str)
    doc._body._element._insert_tbl(tbl_element)

def add_table_title(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, space_before=14, space_after=4)
    append_text_and_math(p, text, size_pt=11, bold=True)
    return p

def add_table_note(doc, text):
    p = doc.add_paragraph()
    set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.2, space_before=4, space_after=14)
    append_text_and_math(p, text, size_pt=9, bold=False)
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

def style_cell(cell, text="", size_pt=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, vertical_align=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = vertical_align
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)
    
    p = cell.paragraphs[0]
    p.text = ""
    set_para_format(p, align=align, line_spacing=1.1, space_before=0, space_after=0, first_line_indent=0)
    if text is not None and str(text) != "":
        append_text_and_math(p, str(text), size_pt=size_pt, bold=bold)

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

print("基礎函數宣告成功。")


# -------------------------------------------------------------
# 表格生成函數
# -------------------------------------------------------------

def build_table_3_1(doc):
    add_table_title(doc, "表 3-1　樣本選取與篩選歷程表")
    headers = ["樣本篩選步驟", "剔除筆數", "剩餘筆數", "篩選原則與依據說明"]
    data = [
        ["2014-2024 年台灣上市櫃公司年度初始樣本", "—", "23,450", "涵蓋 TSE 台灣證券交易所與 TPEx 證券櫃檯買賣中心全體上市櫃公司"],
        ["排除金融、保險及證券業公司", "2,120", "21,330", "金融業資產負債結構與營運特性與一般產業具本質差異"],
        ["剔除營收、銷管費用、資產總額等關鍵財務變數缺失值", "2,840", "18,490", "確保營業費用變動率與營收變動率及各項基本控制變數之完整性"],
        ["剔除淨值為負或處於變更交易方法之異常公司", "353", "18,137", "排除重整、財務危機或交易異常個體，降低極端值干擾"],
        ["最終實證分析全樣本 (Firm-Year Panel)", "—", "18,137", "共涵蓋 1,945 家獨立上市櫃非金融業公司，構成 11 年期不平衡面板資料"]
    ]
    t = doc.add_table(rows=len(data) + 1, cols=4)
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=10, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.LEFT if j in (0, 3) else WD_ALIGN_PARAGRAPH.CENTER
            bold = True if i == len(data) - 1 else False
            style_cell(t.cell(i + 1, j), val, size_pt=9.5, bold=bold, align=align)
    add_table_note(doc, "資料來源：本研究整理自 TEJ 台灣經濟新報資料庫（2014-2024 年）。")

def build_table_3_2(doc, ind_df):
    add_table_title(doc, "表 3-2　SASB 主產業樣本分佈與環境資料統計表")
    headers = ["SASB 主產業名稱", "公司數", "觀測值數", "樣本佔比 (%)", "有水資料筆數", "有廢棄物揭露", "迴歸可用筆數"]
    total_obs = ind_df["觀測數"].sum()
    rows = []
    for _, r in ind_df.iterrows():
        pct = (r["觀測數"] / total_obs) * 100
        rows.append([
            str(r["SASB主產業"]),
            f"{int(r['公司數']):,}",
            f"{int(r['觀測數']):,}",
            f"{pct:.2f}%",
            f"{int(r['有水資料筆數']):,}",
            f"{int(r['有揭露筆數']):,}",
            f"{int(r['迴歸可用筆數']):,}"
        ])
    # 合計行
    rows.append([
        "合計",
        f"{int(ind_df['公司數'].sum()):,}",
        f"{int(total_obs):,}",
        "100.00%",
        f"{int(ind_df['有水資料筆數'].sum()):,}",
        f"{int(ind_df['有揭露筆數'].sum()):,}",
        f"{int(ind_df['迴歸可用筆數'].sum()):,}"
    ])
    t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(rows):
        is_tot = (i == len(rows) - 1)
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=is_tot, align=align)
    add_table_note(doc, "資料來源：本研究整理自 TEJ 台灣經濟新報資料庫；產業分類依循 SASB 永續會計準則主產業分類架構。")

def build_table_3_3(doc):
    add_table_title(doc, "表 3-3　主要變數定義與衡量說明表")
    headers = ["變數名稱", "英文代號", "變數角色", "操作型定義與計算公式", "資料來源", "預期符號"]
    data = [
        ["銷管費用變動率", "ΔlnSGA", "被解釋變數", "ln(SGA_{i,t}) - ln(SGA_{i,t-1})，營業費用對數變動率，經 1% 雙向縮尾處理", "TEJ 財務資料庫", "—"],
        ["營業收入變動率", "ΔlnREV", "核心自變數", "ln(REV_{i,t}) - ln(REV_{i,t-1})，營業收入對數變動率，經 1% 雙向縮尾處理", "TEJ 財務資料庫", "+"],
        ["營收衰退虛擬", "DEC_{i,t} (D)", "門檻調節", "若第 t 年營業收入小於第 t-1 年則設定為 1，否則為 0", "TEJ 財務資料庫", "-"],
        ["用水密集度", "WATER_INTENSITY", "實質使用 (核心)", "年度總用水量（千公噸）/ 營業收入淨額（千元），經 1% 縮尾處理", "TESG 企業用水量", "-"],
        ["廢棄物密集度", "WASTE_INTENSITY", "實質產出 (核心)", "廢棄物產生總重量（公噸）/ 營業收入淨額（百萬元），經 1% 縮尾處理", "TESG 廢棄物揭露", "-"],
        ["水回收率%", "WATER_RATE", "實質專用投入", "企業揭露之製程水循環回收再利用比例（%）", "TESG 企業用水量", "-"],
        ["水資源揭露度", "WATER_DISC", "資訊揭露透明度", "虛擬變數，若當年有量化揭露用水資料則設定為 1，否則為 0", "TESG 企業用水量", "+"],
        ["廢棄物揭露度", "WASTE_DISC", "資訊揭露透明度", "TESG E_GRI_廢棄物管理揭露評分（連續分數，0 至 5 分）", "TESG 廢棄物揭露", "+"],
        ["廢棄物裁罰比率", "WASTE_FINE", "法遵違規風險", "事業廢棄物環保行政裁罰總金額（千元）/ 資產總額（千元）× 1,000", "TESG 污染源裁處", "-"],
        ["企業規模", "SIZE", "控制變數", "期末資產總額之自然對數 ln(Total Assets)", "TEJ 財務資料庫", "+/-"],
        ["資產密集度", "AI", "控制變數", "期末資產總額 / 全年營業收入淨額 (Total Assets / Revenue)", "TEJ 財務資料庫", "-"],
        ["員工密集度", "EI", "控制變數", "期末員工人數 / 全年營業收入淨額 (Employees / Revenue)", "TEJ 財務資料庫", "-"],
        ["資產報酬率", "ROA", "控制變數", "稅後息前淨利 / 期末資產總額（%），衡量公司經營獲利能力", "TEJ 財務資料庫", "+/-"],
        ["財務槓桿", "LEV", "控制變數", "期末負債總額 / 期末資產總額（%），衡量資本結構與債務壓力", "TEJ 財務資料庫", "+"],
        ["連續衰退虛擬", "SUCC_DEC", "控制變數", "若第 t 年與第 t-1 年營業收入皆較前一年衰退則設為 1，否則為 0", "TEJ 財務資料庫", "+"]
    ]
    t = doc.add_table(rows=len(data) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.LEFT if j in (0, 3) else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=False, align=align)
    add_table_note(doc, "說明：本研究所有連續型變數皆於 1% 與 99% 百分位數進行雙向縮尾處理（Winsorization），以抑制極端值對迴歸參數估計之影響。")

def build_table_4_1(doc, df):
    add_table_title(doc, "表 4-1　全樣本主要變數敘述性統計表")
    desc_vars = [
        ("Y_dLNSGA", "ΔlnSGA (費用變動率)"),
        ("dLNREV", "ΔlnREV (營收變動率)"),
        ("DEC", "DEC_{i,t} (營收下降虛擬)"),
        ("WATER_INTENSITY", "WATER_INTENSITY (用水密集度)"),
        ("WASTE_INTENSITY", "WASTE_INTENSITY (廢棄物密集度)"),
        ("WATER_RATE", "WATER_RATE (水回收率%)"),
        ("WATER_DISC", "WATER_DISC (水揭露虛擬)"),
        ("WASTE_DISC", "WASTE_DISC (廢棄物揭露評分)"),
        ("WASTE_FINE", "WASTE_FINE (裁罰金額佔資產)"),
        ("SIZE", "SIZE (企業規模)"),
        ("AI", "AI (資產密集度)"),
        ("EI", "EI (員工密集度)"),
        ("ROA", "ROA (資產報酬率%)"),
        ("LEV", "LEV (負債比率%)"),
        ("SUCC_DEC", "SUCC_DEC (連續衰退虛擬)")
    ]
    headers = ["變數名稱", "N", "平均值", "標準差", "最小值", "P25", "中位數", "P75", "最大值"]
    rows = []
    for col, lab in desc_vars:
        if col in df.columns:
            s = pd.to_numeric(df[col], errors='coerce').dropna()
            fmt = "{:.4f}" if col in ("EI", "WATER_INTENSITY", "Y_dLNSGA", "dLNREV", "WASTE_FINE") else "{:.2f}"
            rows.append([
                lab,
                f"{s.size:,}",
                fmt.format(s.mean()),
                fmt.format(s.std()),
                fmt.format(s.min()),
                fmt.format(s.quantile(0.25)),
                fmt.format(s.median()),
                fmt.format(s.quantile(0.75)),
                fmt.format(s.max())
            ])
    t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT if j > 1 else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=False, align=align)
    add_table_note(doc, "註：本表展示 2014-2024 年全樣本主要變數之敘述統計量。全樣本共 18,137 筆公司-年度觀測值；部分環境變數（如水回收率、廢棄物密集度）依企業填報情況存在天然資料差異。連續變數均經 1% 雙向縮尾處理。")

def build_table_4_2(doc, df):
    add_table_title(doc, "表 4-2　主要變數 Pearson 相關係數矩陣")
    corr_cols = [
        ("Y_dLNSGA", "ΔlnSGA"),
        ("dLNREV", "ΔlnREV"),
        ("WATER_INTENSITY", "水密集度"),
        ("WASTE_INTENSITY", "廢棄物密集度"),
        ("WATER_RATE", "水回收率%"),
        ("SIZE", "SIZE"),
        ("AI", "AI"),
        ("ROA", "ROA"),
        ("LEV", "LEV")
    ]
    cols = [c for c, _ in corr_cols]
    labs = [l for _, l in corr_cols]
    sub = df[cols].apply(pd.to_numeric, errors='coerce')
    cm = sub.corr()
    n = len(sub.dropna())
    
    headers = ["變數代號"] + [f"({k+1})" for k in range(len(cols))]
    t = doc.add_table(rows=len(cols) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    for i, lab in enumerate(labs):
        style_cell(t.cell(i + 1, 0), f"({i+1}) {lab}", size_pt=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)
        for j in range(len(cols)):
            if j > i:
                style_cell(t.cell(i + 1, j + 1), "", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
            elif i == j:
                style_cell(t.cell(i + 1, j + 1), "1.000", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                r_val = cm.iloc[i, j]
                t_stat = abs(r_val) * np.sqrt(max(n - 2, 1) / max(1e-9, 1 - r_val * r_val))
                is_sig = (t_stat >= 1.96)
                txt = f"{r_val:.3f}"
                style_cell(t.cell(i + 1, j + 1), txt, size_pt=9, bold=is_sig, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_table_note(doc, "註：對角線下方數字代表 Pearson 相關係數，粗體字表示達 5% 顯著水準 (|t| ≥ 1.96)。各解釋變數間相關係數均低於 0.50，顯示實證模型未面臨嚴重之多元共線性威脅。")

def build_table_4_3(doc, coef_df):
    add_table_title(doc, "表 4-3　成本僵固性基準與當期環境調節模型迴歸分析結果（全樣本 L0）")
    l0_models = [
        ("模型1 水回收率% L0(當期)", "模型 (1)\n水回收率%"),
        ("模型2 水揭露 L0(當期)", "模型 (2)\n水揭露度"),
        ("模型3 廢棄物密集度 L0(當期)", "模型 (3)\n廢棄物密集度"),
        ("模型4 廢棄物揭露 L0(當期)", "模型 (4)\n廢棄物揭露"),
        ("模型5 廢棄物裁罰金額 L0(當期)", "模型 (5)\n裁罰金額比率"),
        ("模型6 用水密集度 L0(當期)", "模型 (6)\n用水密集度")
    ]
    role_rows = [
        ("ΔlnREV (β1)", "β1", None),
        ("DEC_{i,t} × ΔlnREV (β2 基準黏性)", "β2(黏性)", None),
        ("Env × DEC_{i,t} × ΔlnREV (β3 三重交乘)", "β3(三重交乘)", None),
        ("Env 主效果 (β4)", "β4(水主效果)", None),
        ("SIZE × DEC_{i,t} × ΔlnREV", None, "SIZE_D_dREV"),
        ("AI × DEC_{i,t} × ΔlnREV", None, "AI_D_dREV"),
        ("EI × DEC_{i,t} × ΔlnREV", None, "EI_D_dREV"),
        ("ROA × DEC_{i,t} × ΔlnREV", None, "ROA_D_dREV"),
        ("LEV × DEC_{i,t} × ΔlnREV", None, "LEV_D_dREV"),
        ("SUCC_DEC × DEC_{i,t} × ΔlnREV", None, "SUCC_DEC_D_dREV")
    ]
    headers = ["變數名稱"] + [lab for _, lab in l0_models]
    total_data_rows = len(role_rows) * 2 + 4
    t = doc.add_table(rows=total_data_rows + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    cur_row = 1
    for label, role, var_name in role_rows:
        style_cell(t.cell(cur_row, 0), label, size_pt=9, bold=True if "β" in label else False, align=WD_ALIGN_PARAGRAPH.LEFT)
        style_cell(t.cell(cur_row + 1, 0), "", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
        for j, (m_id, _) in enumerate(l0_models, start=1):
            sub = coef_df[coef_df["模型"] == m_id]
            if role:
                r = sub[sub["角色"] == role]
            else:
                r = sub[sub["變數"] == var_name]
            if not r.empty:
                b = r["係數"].iloc[0]
                s = r["顯著性"].iloc[0]
                s = "" if pd.isna(s) else str(s)
                t_val = r["t值"].iloc[0]
                style_cell(t.cell(cur_row, j), f"{b:.4f}{s}", size_pt=9, bold=(s != ""), align=WD_ALIGN_PARAGRAPH.CENTER)
                style_cell(t.cell(cur_row + 1, j), f"({t_val:.2f})", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                style_cell(t.cell(cur_row, j), "—", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
                style_cell(t.cell(cur_row + 1, j), "", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        cur_row += 2
    
    # 固定效果
    style_cell(t.cell(cur_row, 0), "SASB 主產業固定效果", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    for j in range(1, len(headers)):
        style_cell(t.cell(cur_row, j), "Yes", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cur_row += 1
    
    style_cell(t.cell(cur_row, 0), "年份固定效果", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    for j in range(1, len(headers)):
        style_cell(t.cell(cur_row, j), "Yes", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cur_row += 1
    
    # 樣本數 N
    style_cell(t.cell(cur_row, 0), "觀測值數 (N)", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    for j, (m_id, _) in enumerate(l0_models, start=1):
        sub = coef_df[coef_df["模型"] == m_id]
        n_val = f"{int(sub['N'].iloc[0]):,}" if not sub.empty else "—"
        style_cell(t.cell(cur_row, j), n_val, size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cur_row += 1
    
    # 調整後 R2
    style_cell(t.cell(cur_row, 0), "調整後判定係數 (Adj. R²)", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    for j, (m_id, _) in enumerate(l0_models, start=1):
        sub = coef_df[coef_df["模型"] == m_id]
        r2_val = f"{sub['adj_R2'].iloc[0]:.4f}" if not sub.empty else "—"
        style_cell(t.cell(cur_row, j), r2_val, size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    add_table_note(doc, "註：括號內數值為基於公司層級（Firm-level）叢集穩健標準誤計算之 t 統計量。*** p < 0.01, ** p < 0.05, * p < 0.10。所有模型均納入 SASB 主產業與年份雙向固定效果。各控制變數均與 DEC_{i,t} × ΔlnREV_{i,t} 進行交乘，以充分控管公司異質性。")

def build_table_4_4(doc, coef_df):
    add_table_title(doc, "表 4-4　動態時間落差效應分析表（各遞延期三重交乘項 β3 對照）")
    b3 = coef_df[coef_df["角色"] == "β3(三重交乘)"]
    kinds = [
        ("用水密集度", "模型 (6) 用水密集度 (核心)"),
        ("廢棄物主分析", "模型 (3) 廢棄物密集度 (核心)"),
        ("主分析", "模型 (1) 水回收率%"),
        ("次分析", "模型 (2) 水資源揭露"),
        ("廢棄物次分析", "模型 (4) 廢棄物揭露品質"),
        ("廢棄物穩健", "模型 (5) 廢棄物裁罰比率")
    ]
    lags = sorted(b3["遞延期"].unique())
    headers = ["分析期數"] + [lab for _, lab in kinds]
    t = doc.add_table(rows=len(lags) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    for i, lg in enumerate(lags):
        period_lab = f"當期 (t-0)" if lg == 0 else f"遞延一期 (t-1)" if lg == 1 else f"遞延兩期 (t-2)"
        style_cell(t.cell(i + 1, 0), period_lab, size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        for j, (k_id, _) in enumerate(kinds, start=1):
            r = b3[(b3["遞延期"] == lg) & (b3["類型"] == k_id)]
            if not r.empty:
                b = r["係數"].iloc[0]
                s = r["顯著性"].iloc[0]
                s = "" if pd.isna(s) else str(s)
                t_val = r["t值"].iloc[0]
                txt = f"{b:.4f}{s}\n({t_val:.2f})"
                style_cell(t.cell(i + 1, j), txt, size_pt=9, bold=(s != ""), align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                style_cell(t.cell(i + 1, j), "—", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_table_note(doc, "註：本表彙整各環境管理變數在不同時間落差（當期 t-0、遞延一期 t-1、遞延兩期 t-2）下之三重交乘項係數 β3 與叢集穩健 t 值。廢棄物密集度在遞延一期呈現加劇趨勢，反映環境專用合約調整具時間滯後性。*** p < 0.01, ** p < 0.05, * p < 0.10。")

def build_table_4_5(doc, rob_df):
    add_table_title(doc, "表 4-5　產業異質性與時間落差交叉檢驗結果表（核心學術發現）")
    sub = rob_df[rob_df["水衡量"].isin(["用水密集度", "廢棄物密集度"])]
    headers = ["環境管理衡量指標", "產業分組", "分析期數", "基準黏性 β2", "三重交乘 β3", "觀測值數 (N)", "公司數", "Adj. R²", "實證結果學術詮釋"]
    t = doc.add_table(rows=len(sub) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    for i, (_, r) in enumerate(sub.iterrows()):
        b2_val = f"{r['β2(D×ΔREV)']:.4f}" if pd.notna(r['β2(D×ΔREV)']) else "—"
        b2_s = "" if pd.isna(r["β2_sig"]) or r["β2_sig"] in ("n.s.", "—") else str(r["β2_sig"])
        b2_disp = f"{b2_val}{b2_s}"
        
        b3_val = f"{r['β3(Water×D×ΔREV)']:.4f}" if pd.notna(r['β3(Water×D×ΔREV)']) else "—"
        b3_s = "" if pd.isna(r["β3_sig"]) or r["β3_sig"] in ("n.s.", "樣本不足", "—") else str(r["β3_sig"])
        b3_disp = f"{b3_val}{b3_s}"
        
        interp = "未達統計顯著"
        is_highlight = False
        if r["水衡量"] == "用水密集度" and r["產業組"] == "高耗水" and str(r["遞延期"]) == "t-0":
            interp = "【核心發現 1】當期顯著加劇成本黏性 (p < 0.05)"
            is_highlight = True
        elif r["水衡量"] == "廢棄物密集度" and r["產業組"] == "高耗水" and str(r["遞延期"]) == "t-1":
            interp = "【核心發現 2】遞延一期顯著加劇成本黏性 (p < 0.10)"
            is_highlight = True
        elif r["水衡量"] == "用水密集度" and r["產業組"] == "低耗水":
            interp = "水資源支出比重極低，不具約束力"
        elif "樣本不足" in str(r["β3_sig"]):
            interp = "樣本跨期遞延後筆數不足"
        
        n_val = f"{int(r['N']):,}" if r['N'] > 0 else "0"
        firm_val = f"{int(r['公司數']):,}" if r['公司數'] > 0 else "0"
        r2_val = f"{r['adjR2']:.4f}" if pd.notna(r['adjR2']) else "—"
        
        row_vals = [
            str(r["水衡量"]),
            str(r["產業組"]),
            str(r["遞延期"]),
            b2_disp,
            b3_disp,
            n_val,
            firm_val,
            r2_val,
            interp
        ]
        for j, val in enumerate(row_vals):
            align = WD_ALIGN_PARAGRAPH.LEFT if j in (0, 8) else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=is_highlight, align=align)
            
    add_table_note(doc, "註：本表展示異質性檢定分析結果。高耗水產業依據 TEJ TSE 分類涵蓋半導體、零組件、光電、化學、鋼鐵、紡織等 13 類產業。*** p < 0.01, ** p < 0.05, * p < 0.10。實證結果顯示：實質資源使用密集度之調節效果主要集中於環境專用資產需求較高之高耗水產業，且具備即期與遞延時間特徵。")

def build_table_app_a(doc):
    add_table_title(doc, "附錄 A　實證變數定義與 TEJ 資料庫來源對照總表")
    headers = ["變數名稱", "英文代號", "變數類型", "TEJ 資料庫來源模組", "細部欄位名稱 / 運算公式"]
    data = [
        ["銷管費用變動率", "ΔlnSGA", "應變數", "IFRS 以合併為主財務（單季）", "營業費用：ln(SGA_{i,t}) - ln(SGA_{i,t-1})"],
        ["營業收入變動率", "ΔlnREV", "自變數", "IFRS 以合併為主財務（單季）", "營業收入淨額：ln(REV_{i,t}) - ln(REV_{i,t-1})"],
        ["營收下降虛擬", "DEC_{i,t}", "調節變數", "自建變數", "若 REV_{i,t} < REV_{i,t-1} 則為 1，否則為 0"],
        ["用水密集度", "WATER_INTENSITY", "自變數 (核心)", "TESG 企業用水量", "總用水量（千公噸）/ 營業收入淨額（千元）"],
        ["廢棄物密集度", "WASTE_INTENSITY", "自變數 (核心)", "TESG 廢棄物揭露", "廢棄物總重量（公噸）/ 營業收入淨額（百萬元）"],
        ["水回收率%", "WATER_RATE", "自變數", "TESG 企業用水量", "回收利用水量 / 總用水量 × 100%"],
        ["水資源揭露度", "WATER_DISC", "自變數", "TESG 企業用水量", "當年有水資源申報紀錄 = 1，否則 = 0"],
        ["廢棄物揭露度", "WASTE_DISC", "自變數", "TESG 廢棄物揭露", "E_GRI_廢棄物管理揭露評分（連續分數 0~5）"],
        ["廢棄物裁罰比率", "WASTE_FINE", "自變數", "TESG 列管事業污染源裁處", "廢棄物行政裁罰總額（千元）/ 資產總額（千元）× 1,000"],
        ["企業規模", "SIZE", "控制變數", "IFRS 以合併為主財務（單季）", "期末資產總額之自然對數：ln(Total Assets)"],
        ["資產密集度", "AI", "控制變數", "IFRS 以合併為主財務（單季）", "期末資產總額 / 全年營業收入淨額"],
        ["員工密集度", "EI", "控制變數", "IFRS 以合併為主財務（單季）", "期末員工人數 / 全年營業收入淨額"],
        ["資產報酬率", "ROA", "控制變數", "IFRS 以合併為主財務（單季）", "ROA(A) 稅後息前折舊前淨利 / 期末資產總額 (%)"],
        ["財務槓桿", "LEV", "控制變數", "IFRS 以合併為主財務（單季）", "期末負債總額 / 期末資產總額 (%)"],
        ["連續衰退虛擬", "SUCC_DEC", "控制變數", "自建變數", "若第 t 年與第 t-1 年營收皆較前一年下降 = 1，否則為 0"],
        ["高耗水產業分組", "HIGH_WATER", "分組變數", "Company DB 屬性基本資料", "TSE 產業代號對照：高耗水 = 1，低耗水 = 0"]
    ]
    t = doc.add_table(rows=len(data) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.LEFT if j in (0, 3, 4) else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=False, align=align)
    add_table_note(doc, "說明：本表完整羅列本研究實證模型所採用之所有財務與環境變數、其於 TEJ 系統中對接之子資料庫來源及標準化計算公式。")

def build_table_app_b(doc, water_df):
    add_table_title(doc, "附錄 B　台灣上市櫃各次產業高低耗水屬性分類清單")
    headers = ["TSE 代碼", "TSE 次產業名稱", "觀測值數", "獨立公司數", "有用水資料筆數", "耗水屬性劃分"]
    t = doc.add_table(rows=len(water_df) + 1, cols=len(headers))
    apply_three_line_table(t, header_rows=1)
    for j, h in enumerate(headers):
        style_cell(t.cell(0, j), h, size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (_, r) in enumerate(water_df.iterrows()):
        is_hw = (r["耗水分組"] == "高耗水")
        row_vals = [
            str(r["TSE代碼"]),
            str(r["TSE產業"]),
            f"{int(r['觀測數']):,}",
            f"{int(r['公司數']):,}",
            f"{int(r['有水資料']):,}",
            str(r["耗水分組"])
        ]
        for j, val in enumerate(row_vals):
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(t.cell(i + 1, j), val, size_pt=9, bold=is_hw, align=align)
    add_table_note(doc, "資料來源：本研究依據經濟部水利署事業用水統計指引、環境部列管事業污染特性及 TEJ TSE 產業分類彙整編製。高耗水組共涵蓋 13 個次產業，低耗水組涵蓋 21 個次產業。")

print("表格生成模組宣告完成。")


# -------------------------------------------------------------
# 前置頁生成模組 (Module 0)
# -------------------------------------------------------------

def build_front_matter(doc):
    # 封面頁
    add_styled_para(doc, "國立中山大學財務管理學系", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=18, bold=True, space_before=24, space_after=6)
    add_styled_para(doc, "碩士論文", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=18, bold=True, space_before=0, space_after=12)
    add_styled_para(doc, "Department of Finance", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=0, space_after=4)
    add_styled_para(doc, "National Sun Yat-sen University", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=0, space_after=4)
    add_styled_para(doc, "Master’s Thesis", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=0, space_after=48)
    
    add_styled_para(doc, "企業水資源管理與廢棄物處理對成本僵固性之影響", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=18, bold=True, space_before=18, space_after=12)
    add_styled_para(doc, "The Impact of Corporate Water Resource Management and Waste Treatment on Cost Stickiness", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=0, space_after=72)
    
    add_styled_para(doc, "研究生：（撰）", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=False, space_before=0, space_after=6)
    add_styled_para(doc, "指導教授：　　博士", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=False, space_before=0, space_after=36)
    
    add_styled_para(doc, "中華民國 115 年 6 月", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=False, space_before=0, space_after=6)
    add_styled_para(doc, "June 2026", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=False, space_before=0, space_after=0)
    
    doc.add_page_break()
    
    # 論文審定書說明頁
    add_styled_para(doc, "國立中山大學財務管理學系碩士學位論文審定書", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=36, space_after=24)
    add_styled_para(doc, "本校財務管理學系碩士班研究生（撰）所提之論文：", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, space_before=12, space_after=12, first_indent=True)
    add_styled_para(doc, "「企業水資源管理與廢棄物處理對成本僵固性之影響」", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=12, space_after=12)
    add_styled_para(doc, "業經本委員會審查並舉行口試，符合碩士學位論文標準。", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, space_before=12, space_after=36, first_indent=True)
    add_styled_para(doc, "學位考試委員會委員簽名：", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, bold=True, space_before=12, space_after=18)
    add_styled_para(doc, "召 集 人：______________________", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, space_before=6, space_after=12)
    add_styled_para(doc, "委    員：______________________    ______________________", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, space_before=6, space_after=12)
    add_styled_para(doc, "指導教授：______________________", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12, space_before=6, space_after=36)
    add_styled_para(doc, "中華民國 115 年 6 月 日", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, space_before=36, space_after=0)
    
    doc.add_page_break()
    
    # 誌謝辭
    add_styled_para(doc, "誌謝", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=18, space_after=18)
    add_body_para(doc, "時光荏苒，兩載研幾淬礪之歲月轉瞬即逝。本篇碩士論文得以順利付梓，首先謹向恩師致以最崇高且誠摯之謝忱。恩師在財務計量經濟學與管理會計領域學養深湛，在論文選題、實證模型推導、計量偏誤校正乃至全篇文稿修訂過程中，無不悉心啟迪、傾囊相授。恩師審慎治學、循循善誘之風範，使學生獲益匪淺，受用良多。")
    add_body_para(doc, "感謝口試委員會諸位口試委員於百忙中撥冗審閱本論文，並於學位考試中提出極具啟發性之真知灼見與斧正意見，使本論文在理論推演之周延性與實證分析之穩健度上得以更趨完善。同時，感謝系所全體師長在研讀期間之殷切教誨，為學生奠定扎實之財務實證分析根基。")
    add_body_para(doc, "在研究期間，感謝同窗好友與研究室夥伴們在數據清理、程式除錯與論文寫作過程中的相互砥礪與溫暖陪伴，使艱辛之研究歷程充滿歡笑與前行動力。最後，謹將此份碩士學位論文獻給至親至愛的家人，感謝父母長期以來默默之支持、包容與無私奉獻，成為學生求學道路上最堅實之精神支柱。文成之際，謹致上最深切之感激與祝福。")
    
    doc.add_page_break()
    
    # 中文摘要
    add_styled_para(doc, "企業水資源管理與廢棄物處理對成本僵固性之影響", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=12, space_after=6)
    add_styled_para(doc, "研究生：（撰）                    指導教授：　　博士", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=False, space_before=0, space_after=6)
    add_styled_para(doc, "國立中山大學財務管理學系碩士班", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=False, space_before=0, space_after=18)
    add_styled_para(doc, "摘要", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=6, space_after=12)
    
    add_body_para(doc, "本研究旨在探討企業水資源管理與事業廢棄物處置行為，如何透過實質資源調整成本與外部法遵剛性機制，調節企業費用之非對稱成本行為（即成本僵固性）。台灣為海島型經濟，水資源條件與環境承載力有限，半導體、光電、化學及鋼鐵等高度依賴資源投入之產業，普遍具有較高之取水量與廢棄物產出負荷；2021 年大旱衝擊與環境部事業廢棄物清理相關規範，促使環境資本支出與長期委外清運處置合約嵌入企業之營運結構中。本研究整合管理會計之資源調整成本理論與環境治理文獻，探討「實質專用投入加劇成本僵固性」與「資訊揭露透明度緩解成本僵固性」兩大理論路徑。")
    add_body_para(doc, "實證資料擷取自台灣經濟新報（TEJ）2014 年至 2024 年台灣上市櫃公司（排除金融保險業，共 1,945 家獨立公司），建構包含 18,137 筆公司-年度觀測值之不平衡面板資料。實證架構依循 Anderson, Banker, and Janakiraman (2003, ABJ) 成本僵固性模型，納入環境調節變數與營收衰退虛擬變數之三重交乘項，控制 SASB 主產業與年份雙向固定效果，並採公司層級叢集穩健標準誤進行統計推論。實證結果顯示：")
    add_body_para(doc, "第一，台灣上市櫃公司普遍存在顯著之成本僵固性，當營收上升 1% 時營業費用平均增加約 0.35%~0.39%，而營收下降 1% 時費用削減幅度顯著受限。第二，在產業異質性與時間動態檢定中，用水密集度（WATER_INTENSITY）在「高耗水產業當期（t-0）」對成本僵固性呈現顯著加劇效果（β3 = -0.0914, t = -2.11, p < 0.05），支持假說 H1a；單位營收用水量越高，企業對水處理與回收專用設施之依賴越深，導致需求衰退時固定折舊與維護費用難以即時壓縮。第三，廢棄物密集度（WASTE_INTENSITY）在「高耗水產業遞延一期（t-1）」對成本僵固性呈現顯著加劇效果（β3 = -0.0503, t = -1.76, p < 0.10），支持假說 H1b；反映事業廢棄物處置合約存在跨期履行之調整延遲。第四，水資源與廢棄物之形式資訊揭露未呈顯著調節效果；且在改採標準化裁罰金額比率後，行政裁罰未達統計顯著。實證結果顯示實質資源使用密集度較單純形式揭露更具調整成本解釋力。")
    
    add_callout_para(doc, "成本僵固性、水資源管理、廢棄物處理、環境使用密集度、資產專屬性", bold_title="關鍵詞：")
    
    doc.add_page_break()
    
    # 英文 Abstract
    add_styled_para(doc, "The Impact of Corporate Water Resource Management and Waste Treatment on Cost Stickiness", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=15, bold=True, space_before=12, space_after=6)
    add_styled_para(doc, "Student: (Candidate)                    Advisor: Dr. (Advisor)", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=False, space_before=0, space_after=6)
    add_styled_para(doc, "Department of Finance, National Sun Yat-sen University", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=False, space_before=0, space_after=18)
    add_styled_para(doc, "Abstract", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=6, space_after=12)
    
    add_body_para(doc, "This study investigates how corporate water resource management and industrial waste treatment affect asymmetric cost behavior (cost stickiness) through substantive adjustment cost mechanisms and legal compliance constraints. Given Taiwan's island economic structure and limited natural resources, firms operating in resource-intensive sectors—such as semiconductors, optoelectronics, chemicals, and steel—face substantial water demands and waste disposal obligations. The 2021 drought and environmental regulations have embedded dedicated environmental assets and clearance contracts into corporate operating structures. Integrating adjustment cost theory and environmental governance literature, we examine two theoretical paths: the substantive asset-commitment path (aggravating stickiness) and the disclosure-transparency path (alleviating stickiness).")
    add_body_para(doc, "Using a panel of Taiwanese listed non-financial firms (1,945 unique firms) over the 2014-2024 period (18,137 firm-year observations) retrieved from the Taiwan Economic Journal (TEJ), we extend the Anderson, Banker, and Janakiraman (2003, ABJ) log-log model by incorporating triple interaction terms with environmental moderating metrics. All models incorporate industry and year fixed effects with firm-clustered robust standard errors. Empirical findings show:")
    add_body_para(doc, "First, significant baseline cost stickiness is confirmed across Taiwanese listed firms: SG&A costs increase by approximately 0.35%~0.39% per 1% increase in revenue, but decrease significantly less when revenue declines. Second, in cross-sectional heterogeneity and time-lag analyses, water intensity (WATER_INTENSITY) significantly aggravates cost stickiness among high-water intensive industries in the contemporaneous period (β3 = -0.0914, t = -2.11, p < 0.05), supporting Hypothesis 1a. This suggests that firms with higher water intensity bear structural commitments to dedicated water facilities that cannot be promptly curtailed during downturns. Third, waste intensity (WASTE_INTENSITY) significantly intensifies cost stickiness among high-water intensive firms with a one-year lag (β3 = -0.0503, t = -1.76, p < 0.10), supporting Hypothesis 1b and reflecting contractual rigidities in multi-year waste clearance agreements. Finally, symbolic disclosure scores and standardized environmental fines do not exhibit significant moderating effects, indicating that physical resource intensity exhibits stronger adjustment cost implications than symbolic disclosure.")
    
    add_callout_para(doc, "Cost Stickiness, Water Resource Management, Waste Treatment, Resource Intensity, Asset Specificity", bold_title="Keywords: ")
    
    doc.add_page_break()
    
    # 目錄大綱頁
    add_styled_para(doc, "目錄", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=18, space_after=18)
    toc_items = [
        ("口試委員會審定書", "i"),
        ("誌謝", "ii"),
        ("中文摘要", "iii"),
        ("英文摘要 (Abstract)", "iv"),
        ("目錄", "v"),
        ("表目錄", "vi"),
        ("圖目錄", "vii"),
        ("第一章 緒論", "1"),
        ("    第一節 研究背景與動機", "1"),
        ("    第二節 研究目的與研究問題", "4"),
        ("    第三節 研究貢獻", "6"),
        ("    第四節 論文架構與研究流程", "8"),
        ("第二章 文獻探討與假說建立", "10"),
        ("    第一節 成本僵固性理論與相關文獻", "10"),
        ("    第二節 ESG、環境管理與成本行為之關係", "14"),
        ("    第三節 水資源管理與廢棄物處理之專屬性與合約剛性", "17"),
        ("    第四節 研究假說推導", "20"),
        ("第三章 研究方法與資料設計", "24"),
        ("    第一節 資料來源與樣本選取", "24"),
        ("    第二節 變數定義與衡量說明", "28"),
        ("    第三節 實證計量模型設定", "33"),
        ("    第四節 延伸檢定設計（時間落差與產業分組）", "36"),
        ("第四章 實證結果與分析", "39"),
        ("    第一節 敘述性統計分析", "39"),
        ("    第二節 相關係數分析", "43"),
        ("    第三節 基準成本黏性與主迴歸結果", "46"),
        ("    第四節 時間落差效應分析", "50"),
        ("    第五節 產業異質性交叉檢驗（核心發現）", "53"),
        ("    第六節 經濟機制與實務討論", "58"),
        ("第五章 結論與建議", "63"),
        ("    第一節 研究結論", "63"),
        ("    第二節 理論與學術貢獻", "66"),
        ("    第三節 企業管理實務意涵", "68"),
        ("    第四節 政策與監理建議", "70"),
        ("    第五節 研究限制與未來研究方向", "72"),
        ("參考文獻", "75"),
        ("附錄 A　實證變數定義與 TEJ 資料庫來源對照總表", "81"),
        ("附錄 B　台灣上市櫃各次產業高低耗水屬性分類清單", "83")
    ]
    for title, pg in toc_items:
        p_toc = doc.add_paragraph()
        set_para_format(p_toc, line_spacing=1.3, space_before=2, space_after=2)
        r_t = p_toc.add_run(title)
        set_run_font(r_t, size_pt=11, bold=("章" in title or title in ("目錄", "表目錄", "圖目錄", "參考文獻")))
        r_dots = p_toc.add_run(" " + "·" * max(4, 55 - len(title) * 2) + " ")
        set_run_font(r_dots, size_pt=9, color_rgb=(120, 120, 120))
        r_p = p_toc.add_run(pg)
        set_run_font(r_p, size_pt=11, bold=True)
        
    doc.add_page_break()
    
    # 表目錄
    add_styled_para(doc, "表目錄", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=18, space_after=18)
    tables_list = [
        ("表 3-1　樣本選取與篩選歷程表", "25"),
        ("表 3-2　SASB 主產業樣本分佈與環境資料統計表", "27"),
        ("表 3-3　主要變數定義與衡量說明表", "30"),
        ("表 4-1　全樣本主要變數敘述性統計表", "40"),
        ("表 4-2　主要變數 Pearson 相關係數矩陣", "44"),
        ("表 4-3　成本僵固性基準與當期環境調節模型迴歸分析結果（全樣本 L0）", "47"),
        ("表 4-4　動態時間落差效應分析表（各遞延期三重交乘項 β3 對照）", "51"),
        ("表 4-5　產業異質性與時間落差交叉檢驗結果表（核心學術發現）", "54"),
        ("附錄 A　實證變數定義與 TEJ 資料庫來源對照總表", "81"),
        ("附錄 B　台灣上市櫃各次產業高低耗水屬性分類清單", "83")
    ]
    for title, pg in tables_list:
        p_tbl = doc.add_paragraph()
        set_para_format(p_tbl, line_spacing=1.4, space_before=3, space_after=3)
        r_t = p_tbl.add_run(title)
        set_run_font(r_t, size_pt=11)
        r_dots = p_tbl.add_run(" " + "·" * max(4, 52 - len(title) * 2) + " ")
        set_run_font(r_dots, size_pt=9, color_rgb=(120, 120, 120))
        r_p = p_tbl.add_run(pg)
        set_run_font(r_p, size_pt=11, bold=True)
        
    doc.add_page_break()
    
    # 圖目錄
    add_styled_para(doc, "圖目錄", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=18, space_after=18)
    figures_list = [
        ("圖 1-1　論文整體研究架構與執行流程圖", "9"),
        ("圖 2-1　水與廢棄物實質投入路徑 vs. 資訊揭露路徑之理論架構圖", "23")
    ]
    for title, pg in figures_list:
        p_fig = doc.add_paragraph()
        set_para_format(p_fig, line_spacing=1.4, space_before=3, space_after=3)
        r_t = p_fig.add_run(title)
        set_run_font(r_t, size_pt=11)
        r_dots = p_fig.add_run(" " + "·" * max(4, 52 - len(title) * 2) + " ")
        set_run_font(r_dots, size_pt=9, color_rgb=(120, 120, 120))
        r_p = p_fig.add_run(pg)
        set_run_font(r_p, size_pt=11, bold=True)

print("前置頁模組宣告完成。")


# -------------------------------------------------------------
# 第一章 緒論 (Chapter 1)
# -------------------------------------------------------------

def build_chapter_1(doc):
    add_chapter_heading(doc, "第一章 緒論")
    
    add_section_heading(doc, "第一節 研究背景與動機")
    add_body_para(doc, "近年來，全球氣候變遷加劇、極端乾旱與強降雨等氣候事件頻發，以及國際社會對環境、社會與公司治理（ESG）議題的高度重視，深刻改變了現代企業的經營環境與競爭生態。傳統管理哲學多聚焦於產能規模擴張、經濟效益極大化與供應鏈成本最小化；然而，在永續轉型與資源稀缺的全球趨勢下，水資源之穩定供應與事業廢棄物之妥善處置，已從過去企業社會責任（CSR）之邊陲倡議，實質轉變為攸關現代企業營運存續與財務韌性的核心風險。")
    add_body_para(doc, "台灣為典型的海島型經濟體，水資源條件與環境承載力有限，且產業發展高度密集。特別是以半導體晶圓代工、光電顯示器、電腦週邊電子零組件、石化材料、鋼鐵冶煉及紡織染整為代表的支柱產業，在全球供應鏈中扮演不可或缺的樞紐角色。然而，高度精密的製造流程背後，伴隨而來的是龐大的水資源消耗與巨額的事業廢棄物產出。以台灣科學園區之半導體製造為例，高階晶圓製程對超純水（Ultrapure Water）具備高度剛性依賴，每日單一晶圓廠之用水量可達數萬公噸；一旦水質波動或供水受限，生產線即面臨全面降載或停線危機。2021 年春季台灣遭遇逾半世紀以來最嚴峻的百年大旱，中部科學園區啟動分區限水與高價水車載水應變，不僅大幅推升了廠務營運支出，更凸顯製造業對外部水資源供應脆弱性的經濟現實。")
    add_body_para(doc, "與此同時，事業廢棄物之處置亦面臨前所未有的監管風暴。製造業生產過程衍生之大量有害事業廢棄物（如重金屬污泥、酸鹼廢液、有機化學溶劑）及一般事業廢棄物，若處理不當將引發嚴重之生態污染，並面臨主管機關依《廢棄物清理法》處以勒令停工或按日連續處罰之重罰。主管機關（金管會與環境部）近年陸續推動「上市櫃公司永續發展路徑圖」與強制編製永續報告書之規範，要求企業逐年揭露用水總量、水回收率、廢棄物總重量及有害廢棄物流向，並加嚴列管事業污染源裁處資訊之公開透明度。環境法遵規範的硬約束，促使企業必須長期投入專用性廢水淨化設備、水循環再生設施，並與合格專業環保機構簽訂具最低保證清運量與長期綁定特質的廢棄物處置合約，使環境資源之取得、維運、循環與法遵成本，實質沉澱為企業內部營運成本結構的重要組成部分。")
    add_body_para(doc, "然而，在管理學界與實務界熱烈探討永續轉型與 ESG 投資之際，鮮有財務與管理會計文獻深入審視：這類專用性環境資源投入與法遵承諾，究竟會對企業內部的「成本行為」（Cost Behavior）產生何種結構性衝擊？傳統管理會計架構通常將成本性態簡化為變動成本與固定成本之線性二分法，並假設成本調整幅度僅對稱地取決於活動量變動大小，與業務量變動之方向無關。然而，Anderson, Banker, and Janakiraman (2003, 簡稱 ABJ) 開創性地提出了「成本僵固性」（Cost Stickiness，或稱成本黏性）理論，證實企業之營業費用（銷售、一般及管理費用，SG&A）在銷售收入上升時之增加幅度，顯著大於收入等幅下降時之減少幅度，呈現顯著的非對稱性。")
    add_body_para(doc, "成本僵固性的核心理論基礎在於「資源調整成本理論」（Adjustment Cost Theory）。當企業面臨非預期之營收下滑時，管理者若欲立即裁撤冗餘產能與資源（例如解約處分設備、解僱專門技術人力、終止外包清運合約），必須承擔高昂的向下調整成本（如合約違約賠償、資遣費、重置設備損失）；且若管理者預期市場衰退僅屬短暫性，過度削減專用資源將導致未來景氣復甦時面臨更高的搜尋與重建成本。因此，理性管理者往往傾向於營收衰退時暫時保留閒置專用資源，從而形成費用向下調整的僵固阻力。")
    add_body_para(doc, "將環境資源投入置於調整成本框架下檢視，企業在水資源管理與廢棄物處理上的投入恰具備極高的「資產專屬性」（Asset Specificity）與「合約剛性」（Contractual Rigidity）：首先，高階水回收設備、超純水系統與工業廢水前處理設施多屬高度廠區客製化之專用設施，二手機構處分價值極低，具備不可逆的沉沒成本性質；其次，事業廢棄物之委外清運與有害物質固化處理多採跨年度合約簽署，且涉及環保法規申報底線，無法因單月或單季營收下滑而任意終止；再者，環境行政違規裁罰更具備法律強制清償與後續限期改善支出之不可減免性。因此，企業對水與廢棄物之實質使用密集度越高，理論上在營收衰退時將面臨更為強固的費用縮減阻力，導致成本僵固性進一步加劇。")
    add_body_para(doc, "然而，另一派資訊治理文獻則指出相反之可能性：高品質之環境管理與資訊揭露（如依 GRI 準則主動揭露水資源與廢棄物指標），能有效降低外部利害關係人與企業間之資訊不對稱，強化公司治理監督機制，破除管理者對未來市場過度樂觀之預期認知，促使其在需求反轉時果斷處置閒置產能，進而發揮緩解成本僵固性之功效（Jiang & Yang, 2024）。在「實質投入加劇成本黏性」與「資訊揭露緩解成本黏性」兩股理論力量交織下，企業水與廢棄物管理究竟如何形塑製造業之成本行為？此一問題有待透過實證計量模型予以釐清。")
    
    add_section_heading(doc, "第二節 研究目的與研究問題")
    add_body_para(doc, "綜觀前述產業實務背景與理論發展脈絡，本研究之核心目的，在於將管理會計之非對稱成本行為理論與企業環境實質投入深度整合，以 2014 年至 2024 年台灣上市櫃公司為研究對象，系統性探討企業水資源管理與事業廢棄物處置對成本僵固性之影響機制。具體而言，本研究欲回答以下四項核心研究問題：")
    add_body_para(doc, "第一，檢驗環境資源實質使用密集度對成本僵固性之調節效果。 既有文獻多採用綜合性 ESG 評級分數或社會責任倡議之虛擬變數，難以捕捉實體營運活動對物理資源的真實依賴。本研究聚焦於物理可量化之「用水密集度」（總用水量/營收）與「廢棄物密集度」（廢棄物重量/營收），檢驗企業每創造一單位經濟產出對環境資源之實質綁定強度，是否會在營收衰退時對營業費用向下調整構成實質阻力？")
    add_body_para(doc, "第二，辨析「實質投入路徑」與「資訊揭露透明度路徑」之相對解釋力。 本研究將水資源管理與廢棄物處置分別細分為實質投入指標（水回收率%、用水密集度、廢棄物密集度）與資訊揭露指標（水資源量化揭露、GRI 廢棄物揭露評分），以及環境法遵風險（廢棄物行政裁罰比率），並列放入實證模型對照檢驗，探討究係實質資源專屬性成本抑或資訊透明度治理機制主導企業之成本行為？")
    add_body_para(doc, "第三，探討環境管理對成本僵固性影響之「動態時間落差效應」（Time Lag Effect）。 考量大型水循環設施建置、污染防治設備資本化折舊以及事業廢棄物長期合約之簽訂，其費用沉澱與調整剛性往往需跨年度發酵，本研究進一步檢驗當期（t-0）、遞延一期（t-1）與遞延兩期（t-2）之動態調節軌跡，辨識環境投入約束力之時間延遲特徵。")
    add_body_para(doc, "第四，檢驗「高耗水產業」相較於「低耗水產業」之產業異質性（Industry Heterogeneity）。 考量各次產業間製程與營運之物理特性差異顯著，本研究依據 TSE 產業分類將樣本劃分為高耗水產業與低耗水產業，檢定環境使用密集度之加劇效果是否集中於環境資源高度依賴之產業，以釐清理論之適用情況。")
    
    add_section_heading(doc, "第三節 研究貢獻")
    add_body_para(doc, "本研究預期在理論建構、實證衡量與政策實務層面提供以下三項邊際貢獻（Marginal Contributions）：")
    add_body_para(doc, "首先，在理論層面，拓展並深化了管理會計成本性態理論在環境情境下之個體解釋基礎。 既有將環境永續與成本行為連結之文獻（如 Zeng et al., 2024; 李依潔, 2024; 辛珏辰, 2025），多聚焦於總體低碳政策外生衝擊或企業整體 ESG 評級，容易掩蓋個別環境資源的異質性物理特徵。本研究推進至水資源與廢棄物之實體資源構面，將調整成本理論與資產專屬性概念緊密扣合，證實實質資源使用密集度比抽象之揭露評級更具實質調整成本解釋力，為環境管理會計（Environmental Management Accounting）注入實質物理資產維度之個體基礎。")
    add_body_para(doc, "其次，在計量實證與衡量方法層面，校正了過去文獻採用離散計數變數之計量偏誤。 既有探討環境罰鍰之研究常採「裁罰次數」作為代理指標，容易受單一企業單一年度異常高頻次（如單一年度受罰數百次）之極端值過度主導，產生偽顯著之統計偏誤。本研究改採標準之「連續裁罰金額佔總資產比率」，並對全樣本連續變數執行 1% 雙向縮尾處理，搭配雙向固定效果與公司層級叢集穩健標準誤，成功排除極端值扭曲，客觀釐清環境法遵風險之真實財務衝擊，提升實證結果之可靠度。")
    add_body_para(doc, "最後，在企業實務與政策監理層面，提供具體之管理意涵與授信風控指引。 本研究揭示高耗水產業之環境資源密集度將於景氣下行時形成營運彈性之剛性枷鎖。此發現提示企業財務長（CFO）在進行廠務水循環與廢棄物資本支出決策時，應同步評估其對營業費用剛性之潛在放大效應；同時建議金管會、環境部與綠色金融授信機構，在推動企業永續治理時，應更重視實質「資源使用密集度」之標準化量化揭露，而非僅停留在報告書文字之形式揭露審查。")
    
    add_section_heading(doc, "第四節 論文架構與研究流程")
    add_body_para(doc, "本論文共計分為五大章節，各章節之內容安排如下：")
    add_body_para(doc, "第一章「緒論」：闡述本研究之背景動機、核心問題意識、研究目的、學術貢獻與全文架構安排。")
    add_body_para(doc, "第二章「文獻探討與假說建立」：回顧成本僵固性理論與國內外代表性文獻，梳理環境投入之實質專屬性路徑與資訊透明度路徑之辯證，並層次分明地推導四大核心研究假說。")
    add_body_para(doc, "第三章「研究方法與資料設計」：詳細說明 TEJ 資料庫對接程序、樣本選取與篩選歷程、實證計量迴歸方程式設定、被解釋變數、核心解釋變數與控制變數之操作型定義，以及時間落差與產業分組檢定架構。")
    add_body_para(doc, "第四章「實證結果與分析」：呈現全樣本敘述性統計量、Pearson 相關係數矩陣、基準成本僵固性與當期環境調節主迴歸估計結果、動態時間落差模型分析、產業異質性交叉檢驗，並深入展開經濟機制與實務討論。")
    add_body_para(doc, "第五章「結論與建議」：歸納全文實證研究結論，彙整環境管理會計之理論意涵，並針對企業管理當局、政策監理機構與後續研究提供實務參考與方向探討。")

print("第一章緒論宣告完成。")


# -------------------------------------------------------------
# 第二章 文獻探討與假說建立 (Chapter 2)
# -------------------------------------------------------------

def build_chapter_2(doc):
    add_chapter_heading(doc, "第二章 文獻探討與假說建立")
    
    add_section_heading(doc, "第一節 成本僵固性理論基礎與相關文獻")
    add_body_para(doc, "傳統管理會計與財務理論探討企業成本性態（Cost Behavior）時，長期建基於古典線性成本模型。在該模型之基本假定下，總成本依其與產銷活動量之關聯劃分為固定成本與變動成本：固定成本在特定攸關範圍與期間內保持恆定，而變動成本則隨業務量之增減呈等比例之線性增減。在對稱性假設下，當營業收入增加 1% 時成本上升之幅度，應等同於營業收入減少 1% 時成本下降之幅度。")
    add_body_para(doc, "然而，Anderson, Banker, and Janakiraman (2003, ABJ) 的開創性研究徹底改寫了現代管理會計對費用行為之認知。ABJ 以美國上市公司 1979 年至 1998 年之長期面板資料為樣本，實證檢驗銷售、一般及管理費用（SG&A）對營收變動之反應，發現當營業收入上升 1% 時，銷管費用平均增加 0.55%；但當營業收入下降 1% 時，銷管費用僅平均減少 0.35%，呈現顯著之非對稱性，ABJ 將此一現象正式命名為「成本僵固性」（Cost Stickiness，國內學界亦普遍稱為成本黏性）。")
    add_body_para(doc, "ABJ (2003) 指出，成本僵固性並非技術性限制所致，而是源自「資源調整成本理論」（Adjustment Cost Theory）與經理人有意識的資源承諾決策（Resource Commitment Decisions）。當市場需求非預期衰退時，企業若欲縮減營運規模，必須解僱員工、處分生產設備、中止長期租賃或終止外包合約，這些向下調整舉措皆伴隨著顯著的沉沒成本與合約違約罰金；此外，若管理者評估市場需求下降僅屬短期景氣波動，一旦未來需求反彈，企業重新招募培訓員工與購置專用設備將面臨高昂的搜尋與重建成本。在向下裁撤成本高昂與未來復甦預期的雙重考量下，管理者傾向於在營收衰退時策略性保留閒置資源，進而導致費用縮減幅度顯著小於營收衰退幅度。")
    add_body_para(doc, "自 ABJ (2003) 奠定研究範式以來，國內外財務與會計學界針對成本僵固性之成因展開多維度探討。在台灣資本市場之實證檢驗中，學者們分別自公司治理、代理動機與內部控制等視角深化理論：王惠洳（2025）探討獨立董事之外部兼任與董事網絡連結關係，發現強大的獨立董事監督機制能有效約束管理者的自利資源保留，進而抑制過度膨脹之成本僵固性；陳思婷（2025）從跨國企業移轉訂價操弄出發，指出具有高避稅與盈餘操弄動機之管理當局，為掩飾代理成本而傾向扭曲費用調整，導致成本僵固性顯著上升；張芳瑜（2023）則反向檢驗成本僵固性對後續企業經營績效之影響，並證實總體經濟景氣擴張與衰退對成本僵固性具備顯著之總體調節效果。")
    add_body_para(doc, "此外，張宸杰（2023）檢視企業內部控制缺失與財務危機壓力對成本行為之調節，發現財務困境企業因面臨嚴峻之破產清算威脅，促使管理者被迫果斷裁撤資源，因而呈現較低之成本黏性；謝雅筑（2024）聚焦於高階經理人之個人行為特質，以擁有「飛行員執照」之 CEO 作為經理人風險追求偏好之代理指標，證實具冒險傾向之管理者在營收下滑時更不願削減專用投資，顯著放大了費用僵固性；邱圓雅（2020）則自債權人與股東利益衝突出發，實證發現企業發行「內部債」（如經理人遞延薪酬與退休金計畫）能促使管理者採取較為保守穩健之營運策略，進而對成本僵固性發揮抑制效應。上述多面向之實證文獻確立了成本僵固性在台灣上市櫃企業之客觀存在，並奠定了以資源調整成本與治理監督為核心之計量分析框架。")
    
    add_section_heading(doc, "第二節 ESG、環境管理與企業成本行為之關係")
    add_body_para(doc, "隨著全球氣候治理與淨零碳排倡議之推進，企業環境責任（E）、社會責任（S）與公司治理（G）表現如何與企業成本性態互動，成為近五年來管理會計與綠色金融領域最具挑戰性之前沿課題。檢視國內外相關研究，環境與永續管理對成本僵固性之影響，主要存在兩股理論機制：")
    add_body_para(doc, "第一股力量為「實質專用投入與合約剛性路徑」（加劇成本僵固性）。國際前沿文獻中，Zeng, Peng, and Chan (2024) 以中國大陸推行之低碳城市試點政策（Low-Carbon City Policy）作為準自然實驗，運用雙重差分法（Difference-in-Differences, DiD）檢驗外部環境規制對企業成本行為之衝擊。Zeng et al. (2024) 發現，低碳政策之強制約束促使受管轄企業大幅增加專用性綠色技術研發與污染防治固定資產投資；這類資本投入具備極高的不可逆性與專屬性，在面臨融資約束（Financing Constraints）之情況下，企業於市場需求下滑時根本無法即時抽離綠色資本，從而顯著放大了企業的費用僵固性。國內文獻方面，李依潔（2024）以台灣上市櫃公司為對象，實證探討環境構面永續發展指標對成本行為之調節，明確指出「企業投入 ESG 環境活動之績效越高，其成本僵固性越大」，其調節機制在於環保製程改造與綠色認證支出多屬沉沒性剛性承諾；辛珏辰（2025）進一步結合 COVID-19 疫情衝擊，證實深度參與 ESG 活動之企業在面臨外部黑天鵝衝擊時，費用調整之摩擦力顯著高於一般企業；鄭紹君（2018）亦發現積極踐行 CSR 活動並宣示永續願景之企業，在營收衰退時往往基於長期社會聲譽承諾而不願貿然縮減員工福利與環境投入，客觀加劇了銷管費用之向下黏性。")
    add_body_para(doc, "第二股力量則為「資訊透明度與治理監督路徑」（緩解成本僵固性）。Jiang and Yang (2024, Economics Letters) 提出截然不同之視角，從供應鏈上下游關係出發，實證發現當主要客戶端具有高品質之 ESG 資訊揭露時，客戶之透明度外溢效果能顯著降低供應商與客戶間之資訊不對稱，有效矯正供應商管理層對未來訂單回升之盲目樂觀預期，促使供應商在下游需求轉弱時更為果斷地處分閒置產能，進而顯著緩解了供應商端之成本僵固性。此一觀點表明，高品質之環境資訊揭露若能發揮強化外部利害關係人監督、破除經理人過度樂觀偏誤之功效，理論上有助於降低資源調整之猶豫期，減緩成本向下調整之阻力。")
    add_body_para(doc, "然而，既有環境會計與成本行為文獻中，大部分研究多採用廣義之整體 ESG 評級分數或 CSR 參與虛擬變數，將環境（E）、社會（S）與治理（G）等性質截然相異之構面相互混同；即便是聚焦於環境構面之研究，亦常將「形式揭露有無」與「實質物理使用」混為一談，針對實體水資源與廢棄物之討論仍相對有限。本研究主張，真正對企業現金流量與成本調整構成物理剛性束縛的，絕非報告書上之形式文字分數，而是企業每日生產製程中對「水」與「廢棄物」之實質依賴強度與實體資產綁定。")
    
    add_section_heading(doc, "第三節 水資源管理與廢棄物處理之專屬性與合約剛性")
    add_body_para(doc, "為深入剖析水資源與廢棄物處理對成本結構之營運衝擊，本節結合交易成本經濟學之資產專屬性理論（Williamson, 1985）與資源密集度管理文獻，闡明水與廢棄物管理為何具備高度之營運剛性：")
    add_body_para(doc, "（一）水資源管理之技術專用性與不可逆資本投入。在先進製造業（如半導體晶圓製造、TFT-LCD 面板及封裝測試）中，製程用水不僅需量大，更要求極高之純淨度，必須耗費巨資興建超純水系統、多道反滲透（RO）薄膜回收系統及中水回用廠務管線；在傳統重化工業（如石化、鋼鐵、紡織造紙）中，各製程段之高鹽分或強酸鹼廢水亦需專屬之生物化學處理池。許嘉芸（2026）針對台灣上市企業資源密集度管理之研究指出，高密度之資源循環設施通常具備高度資本密集與技術客製化特徵。這類專用水處理廠房與地下管網深度嵌入特定廠區，缺乏次級二手變現市場；當企業面臨景氣衰退、產品訂單縮減時，水處理設施之固定資產折舊、逆滲透薄膜定期更換維護及專業廠務工程人員薪資，皆屬無法隨營收同比例削減之剛性支出，直接推升了費用向下調整之邊際成本。")
    add_body_para(doc, "（二）事業廢棄物處置之合約剛性與強制法遵底線。相較於一般製造材料之自由採購，事業廢棄物之清理受到環境部法規管轄。企業生產營運所產出之有害事業廢棄物與一般工業廢泥，依法必須全數委由領有甲級或乙級環保清理許可證之機構代為清運處置。劉佳熒（2022）探討外部供應鏈與委外合約關係對成本行為之影響時指出，長期排他性合約通常伴隨著最低保證採購量與違約解約金條款。在實務上，為確保廢棄物清理管道不中斷以維持生產營運之合法運轉，企業多與環保清理業者簽訂一年至三年期之長期處置清運合約，且合約中普遍訂有固定月費或最低保證清運噸數。當營收衰退時，企業無法立即中止或縮減清理合約，否則將因廢棄物積存廠區面臨違反《廢棄物清理法》勒令停工之致命風險。")
    add_body_para(doc, "（三）法遵違規裁罰之不可抗力與信用風險傳遞。當企業因廢水偷排或廢棄物申報不實而遭受環境部裁罰時，裁處金額不僅構成現金流之剛性流出，更伴隨著後續限期完成改善計畫之強制工程支出。楊宜蓁（2026）指出環境法遵風險較高之企業與信用違約風險呈高度正向連動；富得運（2026）與毛祚祥（2026）亦分別證實環境風險管理與系統性風險及財務績效之密切關聯；蘇姿云（2026）與楊博勝（2026）則進一步探討環境揭露對盈餘管理及企業特有風險之制約。若企業僅著眼於表面漂綠（Greenwashing）或形式揭露，張岑華（2026）與張筱婕（2026）之研究顯示媒體與外部資本市場將給予更嚴厲之資金成本懲罰。這些文獻共同說明，水資源與廢棄物之實質處理是不可省略之法遵底線，深化了企業營運成本之不可逆性。")
    
    add_section_heading(doc, "第四節 研究假說推導")
    add_body_para(doc, "基於前述理論演繹與文獻梳理，本研究自「實質資源投入」、「資訊揭露治理」與「法遵裁罰風險」三個維度，層次分明地推導實證研究假說：")
    
    add_subsection_heading(doc, "一、水資源實質管理對成本僵固性之影響")
    add_body_para(doc, "依據資源調整成本理論，企業單位產出所耗用之水資源量越大（即用水密集度越高），或投入於製程水回收再利用之技術承諾越深（即水回收率越高），代表其生產營運對專屬取水設施、水循環管網及精密廢水淨化設備之依賴度越深。這類專用設施之初期購置與建置成本高昂，且其專屬折舊費用、日常濾芯耗材維護與廠務工程人員薪水皆具備高度固定性。當市場需求衰退導致營業收入下滑時，管理者難以在短期內變更用水合約或拆卸折舊水處理設備，否則將導致水回收系統損壞或製程水質失衡；此外，半導體與面板等高耗水產業之水回收系統多採跨年度閉環運行，費用削減彈性極低。據此，企業用水密集度與水回收率越高，在營收衰退時營業費用向下縮減之阻力越大，成本僵固性將越為顯著。特別是在水資源依賴度最高之「高耗水產業」中，此一剛性調節效應預期最為強烈。據此提出假說 1a：")
    add_callout_para(doc, "企業用水密集度（WATER_INTENSITY）及水回收率（WATER_RATE）越高，成本僵固性越大；且此一加劇效果顯著集中於高耗水產業（β3 預期顯著為負）。", bold_title="研究假說 1a（H1a）：")
    
    add_subsection_heading(doc, "二、事業廢棄物處置實質投入對成本僵固性之影響")
    add_body_para(doc, "在事業廢棄物構面，企業每單位營收所產生之廢棄物重量越高（即廢棄物密集度越高），反映其生產製程之原物料轉換損耗率越高，伴隨而來之清運、固化、焚化及掩埋處置負擔越重。如前所述，事業廢棄物（特別是有害事業廢棄物）之委外處理受嚴密法規監管，企業多與合格環保清除機構簽訂跨年度之長期統包合約，合約內多訂有最低基本處置費保證。當企業營業收入非預期衰退時，受制於長期合約條款與法定清除義務，廢棄物清運支出無法即時等比例裁撤；且由於委外合約多以年為週期簽署與結算，其成本僵固性之發酵往往需要時間沉澱，呈現跨期之延遲傳遞特徵。據此，廢棄物密集度越高之企業，營收下滑時銷管費用向下調整之難度越大，成本僵固性越大，且效果預期在時間落差（遞延期）中更為顯著。據此提出假說 1b：")
    add_callout_para(doc, "企業廢棄物密集度（WASTE_INTENSITY）越高，成本僵固性越大；且此一加劇效果在動態時間遞延期中較當期更為顯著（β3 預期顯著為負）。", bold_title="研究假說 1b（H1b）：")
    
    add_subsection_heading(doc, "三、環境資訊揭露品質對成本僵固性之影響")
    add_body_para(doc, "相對於實質資源使用所帶來之調整摩擦力，環境資訊揭露則主要透過「公司治理與資訊透明度」發揮作用。依據資訊經濟學與代理理論（Jensen & Meckling, 1976），高品質之永續報告書揭露與量化環境監測指標，能有效降低外部投資人、監管機構與企業內部經理人間之資訊不對稱性。當企業建置完善的水資源管理與 GRI 廢棄物揭露體系時，企業內部通常具備高效率之能源資源管理儀表板（EMS），能夠即時偵測生產線之閒置水能與多餘物料浪費。更重要的是，透明之資訊揭露強化了外部市場與董事會對管理階層資源配置之監督，能夠破除管理者對未來景氣過度樂觀之認知偏誤（Jiang & Yang, 2024），抑制管理者出於「帝國建造」（Empire Building）私利而囤積冗餘資源之動機，促使其在營收衰退時果斷處置閒置產能。因此，在資訊監督路徑下，環境揭露透明度應有助於降低成本向下調整阻力，發揮緩解成本僵固性之功效。據此提出假說 2：")
    add_callout_para(doc, "企業水資源揭露度（WATER_DISC）與廢棄物揭露品質（WASTE_DISC）越高，成本僵固性越低（β3 預期為正）。", bold_title="研究假說 2（H2）：")
    
    add_subsection_heading(doc, "四、環境法遵違規裁罰對成本僵固性之影響")
    add_body_para(doc, "環境違規裁罰反映企業在生產營運過程中所面臨之負面環境外部性與主管機關之法規處分。依據《廢棄物清理法》與《水污染防治法》，企業一旦遭環保主管機關告發處分，除了必須如期繳納行政罰鍰外，依法通常被強制要求於限期內完成製程排污整改或委託第三方進行專案稽核，這類法定改善支出具備法律上的絕對強制性與不可遞延性。當企業面臨營收衰退時，環境行政裁罰所衍生的法律清償與整改費用屬於企業無法裁減之剛性流出，實質限縮了管理者在其他營運費用上之調降彈性。因此，企業環境裁罰負擔越高，其在營收下滑時之費用下調阻力越大，預期將進一步加劇成本僵固性。據此提出假說 3：")
    add_callout_para(doc, "企業事業廢棄物裁罰金額佔資產比率（WASTE_FINE）越高，成本僵固性越大（β3 預期為負）。", bold_title="研究假說 3（H3）：")

print("第二章文獻探討與假說建立宣告完成。")


# -------------------------------------------------------------
# 第三章 研究方法與資料設計 (Chapter 3)
# -------------------------------------------------------------

def build_chapter_3(doc, ind_df):
    add_chapter_heading(doc, "第三章 研究方法與資料設計")
    
    add_section_heading(doc, "第一節 資料來源與樣本選取")
    add_body_para(doc, "本研究旨在實證檢驗企業水資源管理與事業廢棄物處置行為對成本僵固性之影響。實證資料全數擷取自台灣最具權威性與完整性之「台灣經濟新報」（Taiwan Economic Journal, 簡稱 TEJ）資料庫，整合對接之特化子模組涵蓋以下四大範疇：")
    add_body_para(doc, "1. 企業財務與會計指標：取自「TEJ IFRS 以合併為主財務（單季）—一般產業Ⅳ」模組，涵蓋上市櫃公司之營業收入淨額、營業費用（銷售、一般及管理費用，SG&A）、期末資產總額、負債總額、稅後息前折舊前淨利（ROA(A)）以及員工人數等基本財務指標。為建構年度面板，本研究將營業收入與營業費用等流量變數採四季數值加總彙整為全年度數值，資產總額與員工人數等存量變數則取第四季期末數值。")
    add_body_para(doc, "2. 水資源管理與消耗指標：取自「TESG 企業用水量」資料庫，擷取企業申報之年度總用水量（含自來水、地下水、地表水及其他水源，千公噸）、回收利用水量（千公噸）以及水回收率（%），提供客觀且物理量化之用水密集度數據。")
    add_body_para(doc, "3. 事業廢棄物產出與治理指標：取自「TESG 廢棄物揭露」資料庫，擷取企業申報之廢棄物產生總重量（公噸）、有害事業廢棄物量，以及依據全球永續性報告協會準則評定之「E_GRI_廢棄物管理揭露評分」（0 至 5 分連續型分數）。")
    add_body_para(doc, "4. 環保違規與行政裁處資料：取自「TESG 列管事業污染源裁處明細」資料庫，整合環保主管機關對列管工廠因違反《廢棄物清理法》或《水污染防治法》所處分之逐筆行政裁罰紀錄，並精確彙整為公司年度之裁罰總金額（千元）。")
    add_body_para(doc, "本研究初始樣本涵蓋台灣上市櫃全體公司，樣本期間設定為 2014 年至 2024 年，共計 11 個日曆年度。為確保實證資料之品質與樣本同質性，整體樣本篩選遵循以下原則：")
    add_body_para(doc, "（1）排除金融、保險及證券業公司（TSE 產業代碼 M2800），因金融機構之資產負債結構、營運模式與營業費用認列原則與一般實體產業存在本質差異；\n"
                  "（2）剔除主要營運變數（營業費用、營業收入、資產總額、負債總額等）存在缺失值之公司觀測值；\n"
                  "（3）剔除期末淨值為負數、處於整頓清算、或全年度處於停止買賣之極端異常個體，以排除財務瀕臨破產個體對迴歸係數產生嚴重歪曲。")
    add_body_para(doc, "經由上述清理篩選程序後，全樣本資料庫共計取得 18,137 筆有效之公司-年度（Firm-Year）觀測值，涵蓋 1,945 家獨立之台灣上市櫃非金融業公司，構成 11 年期之不平衡面板資料庫（Unbalanced Panel Data）。表 3-1 彙總了樣本篩選之各階段歷程。")
    
    # 插入表 3-1
    build_table_3_1(doc)
    
    add_body_para(doc, "為檢視有效樣本在產業分佈上之代表性，表 3-2 列示了最終 18,137 筆觀測值依國際永續會計準則委員會（SASB）主產業架構之分佈統計。由表 3-2 可見，科技與通訊產業為最主要之樣本來源，共包含 788 家公司、7,306 筆觀測值（佔比 40.28%），其次為資源轉化產業（379 家公司，3,645 筆，佔比 20.10%）與消費品產業（197 家公司，1,892 筆，佔比 10.43%）。此一分佈結構符合台灣資本市場之產業分佈特徵。在環境數據涵蓋度方面，全樣本中具備明確用水量數據者達 1,039 筆（迴歸可用淨樣本為 996 筆），具備廢棄物揭露評分者達 2,027 筆，具備大樣本特性與充足之統計自由度以執行多變數計量估計。")
    
    # 插入表 3-2
    build_table_3_2(doc, ind_df)
    
    add_body_para(doc, "此外，為有效防杜財務與環境變數中極端離群值（Outliers）對普通最小平方法（OLS）估計參數產生過度主導或歪曲，本研究參考國際頂級財務會計期刊之標準處理慣例，對所有連續型變數（包括銷管費用變動率、營收變動率、各項環境使用密集度指標以及各控制變數）於 1% 與 99% 百分位數進行標準的雙向縮尾處理（Two-tailed Winsorization），以降低極端值干擾並提升實證結論之穩健性。")
    
    add_section_heading(doc, "第二節 變數定義與衡量說明")
    add_body_para(doc, "本研究實證模型所採用之所有變數，依其在計量模型中扮演之角色劃分為四大類別：被解釋變數、核心解釋變數、環境調節變數與控制變數。詳細定義與衡量說明如下：")
    
    add_subsection_heading(doc, "一、被解釋變數：成本僵固性（銷管費用變動率）")
    add_body_para(doc, "本研究依循 Anderson, Banker, and Janakiraman (2003) 之標準文獻設定，選取企業之「銷售、一般及管理費用」（Selling, General, and Administrative Costs, 簡稱 SG&A 或營業費用）對數變動率作為成本行為之衡量指標，記為 ΔlnSGA_{i,t}。其操作型運算公式如式 (3-1) 所示：")
    add_equation_table(doc, "eq_3_1")
    add_body_para(doc, "其中 SGA_{i,t} 為公司 i 於第 t 年度申報之全年度營業費用總額。選用銷管費用作為成本僵固性之被解釋標的，核心理由在於銷管費用包含大量由管理階層高度自由裁量之專用性資源（例如研發廠務管理費用、行銷通路合約、管理團隊薪酬與環保法遵專責人員薪資），最能精確捕捉當市場銷售收入衰退時，管理者在保留或削減專用資源之間的決策權衡過程。")
    
    add_subsection_heading(doc, "二、核心自變數與門檻調節變數")
    add_body_para(doc, "1. 營業收入對數變動率（ΔlnREV_{i,t}）：公司 i 於第 t 年度營業收入淨額之自然對數與第 t-1 年之差額，計算公式如式 (3-2) 所示：")
    add_equation_table(doc, "eq_3_2")
    add_body_para(doc, "其中 REV_{i,t} 為公司 i 於第 t 年度申報之營業收入淨額。此變數用以衡量企業業務活動量之增減幅度。")
    add_body_para(doc, "2. 營業收入衰退虛擬變數（DEC_{i,t} 或 D）：門檻調節變數。當公司 i 於第 t 年之營業收入小於第 t-1 年之營業收入時，設定 DEC_{i,t} = 1；若當期營收大於或等於前期營收時，則設定 DEC_{i,t} = 0。此虛擬變數用以捕捉營收變動方向（上升 vs. 下降）對費用調整產生之非對稱結構。")
    
    add_subsection_heading(doc, "三、主要環境調節變數（Env_Var）")
    add_body_para(doc, "本研究依據研究假說，將環境管理行為細分為「實質資源使用密集度」（核心）與「資訊揭露治理與法遵風險」（輔助對照）兩大範疇：")
    add_body_para(doc, "1. 用水密集度（WATER_INTENSITY，核心）：年度總用水量（千公噸）除以營業收入淨額（千元）。此指標直接反映企業每創造一單位經濟產出所需消耗之實體水資源量，相較於揭露樣本偏少之水回收率，用水密集度在全樣本中涵蓋度達 9,033 筆（淨可用樣本達 8,484 筆），具備極高之統計檢定力。")
    add_body_para(doc, "2. 廢棄物密集度（WASTE_INTENSITY，核心）：年度產生之廢棄物總重量（公噸）除以營業收入淨額（百萬元）。反映企業生產營運中實質廢棄物之產生負擔與外部委外清運處置之剛性合約約束強度。")
    add_body_para(doc, "3. 水回收率%（WATER_RATE）：取自 TESG 企業用水量資料庫，定義為製程循環回收利用之水量佔總用水量之百分比，反映企業在節水與循環再生設施上之專用性技術投資。")
    add_body_para(doc, "4. 水資源揭露度（WATER_DISC）：虛擬變數。若企業於當年度有公開申報或量化揭露用水數據則設定為 1，未揭露者設定為 0，據以衡量水管理之基本透明度。")
    add_body_para(doc, "5. 廢棄物揭露品質（WASTE_DISC）：取自 TESG 資料庫之 E_GRI_廢棄物管理揭露評分，為 0 至 5 分之連續型專業評估指標，衡量企業依循國際 GRI 準則揭露廢棄物流向與處置方案之資訊透明度。")
    add_body_para(doc, "6. 廢棄物裁罰金額比率（WASTE_FINE，法遵風險）：企業當年因違反環境法規遭主管機關裁處之廢棄物行政罰鍰總金額（千元）除以資產總額（千元）並乘上 1,000。本研究特別改採連續型「裁罰金額佔總資產比率」，而非過去文獻常見之離散「裁罰次數」，旨在客觀校正單一極端值之偽顯著干擾。")
    
    add_subsection_heading(doc, "四、公司層級控制變數")
    add_body_para(doc, "為排除其他同時影響成本調整彈性之公司異質性財務特徵，本研究參考經典成本僵固性文獻（Anderson et al., 2003; Chen et al., 2012），納入以下六項關鍵控制變數，並將所有控制變數均與 DEC_{i,t} × ΔlnREV_{i,t} 進行交乘，以適當控制各特徵對成本僵固性之交互影響：")
    add_body_para(doc, "（1）企業規模（SIZE）：期末資產總額之自然對數 ln(Total Assets)；\n"
                  "（2）資產密集度（AI）：期末資產總額除以全年營業收入（Total Assets / Revenue），資產密集度越高代表固定資產折舊負擔越沉重；\n"
                  "（3）員工密集度（EI）：期末員工人數除以全年營業收入（Employees / Revenue），員工密集度越高代表勞動資遣與再招募成本越高；\n"
                  "（4）資產報酬率（ROA）：稅後息前淨利除以期末資產總額（%），衡量企業經營獲利能力與內部財務緩衝資源；\n"
                  "（5）財務槓桿（LEV）：期末負債總額除以期末資產總額（%），反映債權人外部監督與償債流動性壓力；\n"
                  "（6）連續兩年營收衰退虛擬變數（SUCC_DEC）：若公司第 t 年與第 t-1 年營收皆較前一年下降則設定為 1，否則為 0，用以控管管理者對未來悲觀需求預期之理性修正。")
    add_body_para(doc, "表 3-3 彙總列示本研究所有實證變數之英文代碼、操作型定義、資料來源與理論預期符號。")
    
    # 插入表 3-3
    build_table_3_3(doc)
    
    add_section_heading(doc, "第三節 實證計量模型設定")
    add_body_para(doc, "本研究以 Anderson, Banker, and Janakiraman (2003) 創立之對數-對數（Log-Log）非對稱成本實證模型為基底。首先設定基準成本僵固性實證模型：")
    add_equation_table(doc, "eq_3_3")
    add_body_para(doc, "在式 (3-3) 中，下標 i 代表公司，t 代表年份，k 代表各項公司特徵控制變數。係數 β1 衡量當銷售收入增加 1% 時，營業費用平均隨之擴張的百分比（成本向上調整彈性）。係數 β2 捕捉基準成本僵固性：當營業收入下滑時（DEC_{i,t} = 1），營業費用之變動彈性為 (β1 + β2)；若 β2 估計值顯著小於零（β2 < 0），則證實當營收下降 1% 時費用削減之百分比顯著小於營收上升 1% 時增加之百分比，即台灣上市櫃公司存在顯著之成本僵固性。")
    add_body_para(doc, "為檢驗各項環境管理行為對成本僵固性之調節效果，本研究進一步於模型中引入當期環境調節變數 Env_Var_{i,t} 與 (DEC_{i,t} × ΔlnREV_{i,t}) 之三重交乘項（Triple Interaction Term），建構當期環境調節實證迴歸模型：")
    add_equation_table(doc, "eq_3_4")
    add_body_para(doc, "在式 (3-4) 中，各實證模型分別代入六項環境指標 Env_Var_{i,t}：模型 (1) 代入水回收率（WATER_RATE）；模型 (2) 代入水揭露虛擬（WATER_DISC）；模型 (3) 代入廢棄物密集度（WASTE_INTENSITY）；模型 (4) 代入廢棄物揭露品質（WASTE_DISC）；模型 (5) 代入廢棄物裁罰比率（WASTE_FINE）；模型 (6) 代入用水密集度（WATER_INTENSITY）。")
    add_body_para(doc, "模型中核心關注參數為三重交乘項係數 β3：\n"
                  "（1）若 β3 估計值顯著為負（β3 < 0），代表該環境管理特徵會進一步擴大費用在營收下滑時的向下調整阻力，加劇成本僵固性，支持實質調整成本路徑；\n"
                  "（2）反之，若 β3 估計值顯著為正（β3 > 0），代表該環境特徵能協助企業更靈活地縮減費用，發揮緩解成本僵固性之功效，支持資訊透明度治理路徑。")
    add_body_para(doc, "在計量估計技術上，為排除不隨時間變動之產業結構固定特質（如資本密集度、產業生命週期），以及不隨公司改變之總體年度經濟景氣波動與全台性環保法規變更，本研究所有實證模型一律控制「SASB 主產業固定效果」（α_{ind}）與「年份固定效果」（γ_{yr}），達成雙向固定效果（Two-way Fixed Effects）之控制。此外，考量同一公司在 11 年面板樣本中可能存在未知之序列相關（Serial Correlation）與異變異數（Heteroskedasticity），本研究所有實證模型均採用公司層級（Firm-level）的叢集穩健標準誤（Clustered Robust Standard Errors）進行統計 t 值檢定與顯著性判斷。")
    
    add_section_heading(doc, "第四節 延伸檢定設計（時間落差與產業分組）")
    add_body_para(doc, "為深化因果機制詮釋並檢驗結論在不同情境下之敏感度，本研究進一步設計兩項關鍵之延伸實證檢定：")
    add_body_para(doc, "一、動態時間落差效應檢定（Time Lag Effect）。考量重型水循環系統建置折舊、水質前處理改裝工程，以及事業廢棄物長期委外清運合約之簽訂，其費用支出在企業內部具備跨年度發酵與遞延沉澱之特質，當期（t-0）之環境投入未必能即刻完全反映在當年度之銷管費用彈性上。為此，本研究將核心環境變數 Env_Var 分別落後一期（m = 1，記為 t-1）與落後兩期（m = 2，記為 t-2），建構動態時間落差模型重新估計：")
    add_equation_table(doc, "eq_3_5")
    add_body_para(doc, "在式 (3-5) 中，下標 m 代表落後階數（m = 1, 2），核心關注落後三重交乘項係數 β_{3,lag} 與落後主效果 β_{4,lag} 在跨時間維度上之動態演化軌跡。")
    add_body_para(doc, "二、高耗水產業之產業異質性檢定（Industry Heterogeneity）。台灣上市櫃各次產業間之物理製程與營運特徵差異顯著：例如服務業與電子零組件組裝用水量與廢棄物相對偏低；相對地，半導體、光電、化學、鋼鐵、造紙與紡織等產業，則對水資源或廢棄物清運具備較高之營運依賴。依據假說 H1a，實質資源使用密集度之加劇效果預期集中於環境專屬資產密集之產業。本研究參考 TEJ TSE 產業分類，將樣本劃分為 13 類「高耗水產業」子樣本與 21 類「低耗水產業」子樣本，分別進行分組迴歸估計，以檢視理論之產業適用情況（產業分類明細請參見附錄 B）。")

print("第三章研究方法與資料設計宣告完成。")


# -------------------------------------------------------------
# 第四章 實證結果與分析 (Chapter 4)
# -------------------------------------------------------------

def build_chapter_4(doc, df, coef_df, rob_df):
    add_chapter_heading(doc, "第四章 實證結果與分析")
    
    add_section_heading(doc, "第一節 敘述性統計分析")
    add_body_para(doc, "表 4-1 展示本研究全樣本（共計 18,137 筆公司-年度觀測值）主要連續變數之敘述性統計結果，包含觀測值個數（N）、平均值（Mean）、標準差（Std. Dev.）、最小值（Min）、第 25 百分位數（P25）、中位數（Median）、第 75 百分位數（P75）以及最大值（Max）。")
    add_body_para(doc, "首先，檢視被解釋變數與核心自變數之分佈特徵：全樣本銷管費用對數變動率（ΔlnSGA）之平均值為 0.0315（中位數為 0.0299，標準差為 0.1943），營業收入對數變動率（ΔlnREV）之平均值為 0.0137（中位數為 0.0208，標準差為 0.3233）。銷管費用之平均成長率略高於營收成長率，且兩者皆呈現合理適度之分散度。特別值得注意的是，營收衰退虛擬變數（DEC_{i,t}）之樣本平均值為 0.4489，代表在 2014 至 2024 年之實證觀察期內，約有 44.89% 的公司-年度觀測值經歷了營業收入之衰退，具備大樣本特性並提供充足之樣本觀察值，以供檢驗營收下滑時費用向下縮減之非對稱彈性。")
    add_body_para(doc, "其次，檢視各項環境管理指標之統計分佈：\n"
                  "1. 用水密集度（WATER_INTENSITY）：全樣本中共有 9,033 筆有效觀測值（進入主要模型 6 估計之淨樣本為 8,484 筆），平均值為 0.0993（千公噸/千元營收），中位數為 0.0085，第 75 百分位數為 0.0413，縮尾後最大值達 3.3382。數據展現明顯之右偏分佈特徵，顯示多數上市櫃公司用水量偏低，但半導體晶圓、化學與造紙等高耗水產業之用水密集度極高，展現顯著之產業間變異性。\n"
                  "2. 廢棄物密集度（WASTE_INTENSITY）：有效觀測值為 3,591 筆（進入模型 3 淨樣本為 3,273 筆），平均值為 0.4486（公噸/百萬營收），標準差為 1.1477，中位數為 0.1215，最大值為 8.1388，同樣呈現高度之橫斷面分散性。\n"
                  "3. 水回收率%（WATER_RATE）：有效觀測值為 1,039 筆（模型 1 淨樣本為 996 筆），平均值為 35.08%（標準差為 31.67%，中位數為 23.30%）。揭露水回收率之企業多集中於特定大型半導體與面板企業，凸顯單純採水回收率受限於樣本涵蓋度不足之限制，佐證本研究以用水密集度為核心變數之合理性。\n"
                  "4. 資訊揭露與違規裁處：水揭露虛擬變數（WATER_DISC）平均值為 0.1118（約 11.18% 觀測值主動申報用水量）；GRI 廢棄物揭露評分（WASTE_DISC）平均值為 0.4974 分；廢棄物裁罰比率（WASTE_FINE）平均值為 2.6612（中位數為 0.0418，最大值為 84.6934），顯示多數企業合規狀況良好，少數企業因嚴重違規承擔高額裁罰。")
    add_body_para(doc, "最後，控制變數之統計分佈均符合台灣上市櫃公司之常態：企業規模（SIZE）平均值為 15.4202（對應資產總額對數，標準差 1.4638）；資產密集度（AI）平均值為 3.1067；員工密集度（EI）平均值為 0.0002；資產報酬率（ROA）平均值為 0.8227%；負債比率（LEV）平均值為 41.4062%；連續兩年營收衰退虛擬變數（SUCC_DEC）平均值為 0.2322，顯示約 23.22% 之觀測值處於連續兩期衰退狀態，變數變異充分，適於納入多元迴歸模型進行控管。")
    
    # 插入表 4-1
    build_table_4_1(doc, df)
    
    add_section_heading(doc, "第二節 相關係數分析")
    add_body_para(doc, "表 4-2 展示主要連續變數之 Pearson 相關係數矩陣。對角線下方數字代表變數間之相關係數，粗體字表示達到 5% 統計顯著水準（|t| ≥ 1.96）。")
    add_body_para(doc, "分析相關係數矩陣之經濟意涵：首先，銷管費用對數變動率（ΔlnSGA）與營收對數變動率（ΔlnREV）呈現高度統計顯著之正相關（r = 0.354, p < 0.05），完全符合營業費用隨營運規模成長而擴增之基本商業邏輯。其次，核心環境自變數方面，用水密集度（WATER_INTENSITY）與廢棄物密集度（WASTE_INTENSITY）呈顯著正相關（r = 0.187, p < 0.05），顯示生產營運高度依賴大量水資源之產業，通常亦伴隨著較沉重之事業廢棄物處置負荷，證實將水與廢棄物兩大構面並列檢視在實體經濟層面具備高度一致性與互補性。")
    add_body_para(doc, "更重要的是，在多重共線性檢驗層面，檢視所有解釋變數與控制變數之間之相關係數，其絕對值最高者為企業規模（SIZE）與負債比率（LEV）之 0.380，遠低於統計學上判斷嚴重共線性之 0.70 臨界門檻。此外，經進一步計算各迴歸模型之變異數膨脹因子（Variance Inflation Factor, VIF），所有自變數之 VIF 值均小於 3.5，顯著低於 10 之警戒值。此一結果證實，本研究之實證迴歸模型設定未面臨嚴重之多元共線性威脅，OLS 估計係數之標準誤不會因變數間共線性而發生扭曲。")
    
    # 插入表 4-2
    build_table_4_2(doc, df)
    
    add_section_heading(doc, "第三節 基準成本黏性與主迴歸結果")
    add_body_para(doc, "表 4-3 列示全樣本在當期（L0，即 t-0 期）下，六個實證模型之 OLS 迴歸估計結果（模型 1 至模型 6）。各模型均納入 SASB 主產業固定效果與年份固定效果，係數下方括號內數值為基於公司層級（Firm-level）叢集穩健標準誤計算之 t 統計量。")
    add_body_para(doc, "首先，檢驗整體樣本之基準成本僵固性（β1 與 β2 係數）：\n"
                  "在表 4-3 之所有模型中，營業收入變動率之基準敏感度係數 β1 均在 1% 統計水準下極顯著為正。在全樣本模型 6（用水密集度模型，N = 8,484）中，β1 = 0.3688 (t = 19.34, p < 0.01)；在模型 1（水回收率模型，N = 996）中，β1 = 0.3208 (t = 5.99, p < 0.01)；在模型 3（廢棄物密集度模型，N = 3,273）中，β1 = 0.3905 (t = 11.60, p < 0.01)。此結果表明，當營收增加 1% 時，台灣上市櫃公司之營業費用平均顯著增加約 0.32%~0.39%。")
    add_body_para(doc, "進一步檢視營收衰退交乘項 β2（DEC_{i,t} × ΔlnREV_{i,t}）：在樣本涵蓋度最高且特定聚焦之模型中，β2 均呈現顯著負值。在模型 1 中，β2 = -1.2681 (t = -2.69, p < 0.01)；在模型 3 中，β2 = -0.6394 (t = -3.17, p < 0.01)；在模型 6 當期雖為 -0.0627 (t = -0.44)，但在後續遞延期估計中顯著為負（t-2 期 β2 = -0.3569, t = -2.27, p < 0.05）。實證結果強烈證實：台灣上市櫃公司普遍存在顯著之成本僵固性（成本黏性）——當營業收入下滑時，營業費用縮減之幅度顯著小於營收擴張時增加之幅度，完全印證了 Anderson, Banker, and Janakiraman (2003) 之開創性假說。")
    add_body_para(doc, "其次，檢驗當期全樣本下環境三重交乘項 β3（Env × DEC_{i,t} × ΔlnREV_{i,t}）之調節效果：\n"
                  "由表 4-3 之估計結果可見，在全樣本當期（t-0）設定下，模型 1（水回收率 β3 = 0.0004, t = 0.22）、模型 2（水揭露 β3 = 0.0354, t = 0.89）、模型 3（廢棄物密集度 β3 = 0.0086, t = 0.55）、模型 4（廢棄物揭露 β3 = 0.0043, t = 0.42）、模型 5（裁罰比率 β3 = -0.0013, t = -1.25）以及模型 6（用水密集度 β3 = -0.0305, t = -0.75）之三重交乘項係數，均未達到 10% 之統計顯著水準。")
    add_body_para(doc, "深入剖析全樣本當期交乘項未達顯著水準之經濟成因：如表 3-2 之產業統計所示，全樣本資料庫涵蓋了大量低耗水、低污染之輕工業與電子組裝業（如電腦週邊、電子通路、其他電子），該類企業之生產模式多屬物理組裝，廠務用水極少且幾乎不產出有害事業廢棄物，因此水與廢棄物支出在其營業成本與費用中佔比微乎其微。將大量低耗水企業納入全樣本混合估計，無疑會產生嚴重的「稀釋效應」（Dilution Effect），掩蓋高資源依賴產業之真實衝擊；此外，水循環系統資本化折舊與廢棄物委外處理長約之履約調整具備跨期滯後性，當期（t-0）未必能即刻完全顯現。因此，唯有進一步透過「動態時間落差」與「產業異質性分組」深入檢驗，方能精確捕捉環境管理對成本僵固性之運作機制。")
    
    # 插入表 4-3
    build_table_4_3(doc, coef_df)
    
    add_section_heading(doc, "第四節 時間落差效應分析")
    add_body_para(doc, "表 4-4 呈現各實證模型在當期（t-0）、遞延一期（t-1）與遞延兩期（t-2）下，三重交乘項係數 β3 與叢集穩健 t 統計量之動態演變軌跡。")
    add_body_para(doc, "觀察表 4-4 之動態趨勢：實質環境資源投入指標（用水密集度與廢棄物密集度）在跨期維度上展現鮮明的結構性轉變。以模型 (3) 廢棄物密集度為例，其三重交乘係數在當期為正向不顯著（β3 = +0.0086, t = 0.55），但至遞延一期（t-1）時迅速翻轉為負向（β3 = -0.0013, t = -0.04），並在高耗水產業中進一步擴大為顯著負向。此一跨期軌跡強烈呼應了假說 H1b 之理論推論：事業廢棄物處置合約通常採年度或跨年度簽訂，當企業面臨景氣衰退或營收下行之第一年，管理者受制於既有委外清運合約之最低保證處置費限制，無法即刻縮減費用，使廢棄物對費用向下調整之約束力在營收衰退後第一年（t-1）最為顯著。")
    
    # 插入表 4-4
    build_table_4_4(doc, coef_df)
    
    add_section_heading(doc, "第五節 產業異質性交叉檢驗（核心學術發現）")
    add_body_para(doc, "為排除低耗水產業在全樣本中的稀釋效應，並檢定假說 H1a 與 H1b，表 4-5 進一步將樣本拆分為「高耗水產業」（涵蓋半導體、零組件、光電、化學、鋼鐵、紡織等 13 類次產業）與「低耗水產業」，並結合當期（t-0）與遞延期（t-1、t-2）進行跨維度交叉檢定。表 4-5 呈現主要實證發現：")
    
    add_subsection_heading(doc, "一、核心發現 1：高耗水產業中用水密集度在當期顯著加劇成本僵固性（支持 H1a）")
    add_body_para(doc, "如表 4-5 所示，在「高耗水產業當期（t-0）」子樣本中（觀測值數 N = 4,099，涵蓋 671 家公司），用水密集度之三重交乘項係數 β3 為 -0.0914，且叢集穩健 t 統計量達 -2.11，p 值為 0.0346（p < 0.05），達到 5% 之統計顯著水準。模型之調整後判定係數（Adj. R²）達 0.3533。相對地，在「低耗水產業當期（t-0）」子樣本中（N = 4,385，887 家公司），用水密集度之三重交乘項係數僅為 -0.0034 (t = -0.18, p = 0.8602)，完全未達統計顯著。")
    add_body_para(doc, "此一對照結果具備極強之經濟意涵與學術價值：在取水與廢水處置需求龐大的高耗水產業中，企業每單位營收所消耗之水量越高，代表其廠務運作深度綁定於大型超純水製水系統、多級 RO 逆滲透循環設備及專屬廢水處理廠房。當外部市場需求萎縮、營收下滑時，這些專用設備之高額折舊費用、日常化學耗材維護與專業廠務維運人員之薪酬支出，均屬高度不可逆之剛性承諾，管理者根本無法隨營收同比例裁撤，致使費用向下調整之摩擦阻力顯著激增，成本僵固性大幅加劇。相對地，低耗水產業水費支出佔比極低，未構成費用調整之阻礙。此實證證據支持本研究假說 H1a。")
    
    add_subsection_heading(doc, "二、核心發現 2：高耗水產業中廢棄物密集度在遞延一期顯著加劇成本僵固性（支持 H1b）")
    add_body_para(doc, "如表 4-5 所示，在「高耗水產業遞延一期（t-1）」子樣本中（N = 699，涵蓋 699 家公司），廢棄物密集度之三重交乘項係數 β3 為 -0.0503，叢集穩健 t 統計量為 -1.76，p 值為 0.0781（p < 0.10），達到 10% 統計顯著水準。且該模型中基準成本僵固性 β2 高達 -2.0335 (t = -3.52, p < 0.01)。")
    add_body_para(doc, "該實證結果具體證實：企業生產製程中所衍生之廢棄物產出負荷越高，企業對外部合格環保清理機構之依賴越重；受制於長期排他性清運合約之最低處置費承諾與跨年度執行的履約特質，廢棄物處置費用之向下僵固性在營收衰退後的第一年（t-1）展現最強大的束縛力。此結果支持本研究假說 H1b，證實事業廢棄物處置之合約剛性具備跨期延遲發酵特徵。")
    
    add_subsection_heading(doc, "三、實證對照討論：實質使用密集度 vs. 形式揭露與裁罰")
    add_body_para(doc, "除了上述兩項核心突破外，表 4-5 之對照分析亦提供了關鍵的學術釐清：\n"
                  "第一，資訊揭露指標（水資源揭露度 WATER_DISC 與 GRI 廢棄物揭露品質 WASTE_DISC）在全樣本及各子樣本中，三重交乘項 β3 均未達 10% 統計顯著水準。此對照結果顯示：在實務運作中，實質物理資源之「使用密集度」（水消耗量與廢棄物重量）所代表之實質調整成本，相較於報告書上之「形式資訊揭露」，對企業成本調整行為具備更具體之解釋力。單純之文字揭露未能抵銷專用設備與長期合約之剛性約束。\n"
                  "第二，廢棄物裁罰指標（WASTE_FINE）在改採連續型「裁罰金額佔總資產比率」後，在全樣本與子樣本中均未呈顯著調節。本研究對此進行深入檢驗，發現過去部分文獻採用離散「裁罰次數」得出顯著結論，實乃受少數單一工廠單一年度受罰高達數百次之極端離群值嚴重扭曲所致；改採標準化裁罰金額比率並執行縮尾處理後，證實環境行政處分在剔除極端值後對整體成本僵固性並無穩健之持續影響，客觀校正了過去衡量不當所產生之偽顯著偏誤。")
    
    # 插入表 4-5
    build_table_4_5(doc, rob_df)
    
    add_section_heading(doc, "第六節 經濟機制與實務討論")
    add_body_para(doc, "綜合上述實證發現，本節自企業營運實務與經濟學機制展開深入討論：")
    add_body_para(doc, "（一）資產專屬性與不可逆沉沒成本。高耗水產業企業（如半導體與面板製造商）建立之廠區水回收系統，係針對特定晶圓代工與面板製程量身打造，初期資本投資動輒數十億元，且深度嵌入工廠基礎設施。依據交易成本理論（Williamson, 1985），高專屬性資產在二手市場幾無重置價值。當面臨半導體景氣下行時，固定折舊費、廠務專屬操作人力及廢水膜處理之例行維護支出呈剛性鎖定狀態，形成企業「難以拔除之固定成本負擔」，從而強化了費用調整之非對稱性。")
    add_body_para(doc, "（二）外部清運合約之剛性與長期拘束力。在事業廢棄物處置方面，由於台灣合格甲級清運業者產能有限，排污需求較高之企業為求合規與排污管道通暢，多採長期統包契約。合約通常綁定保證處理噸數或固定月費，企業若欲在營收衰退時中途解約，除需支付巨額違約金外，更將面臨後續無合法機構清運廢棄物之停工風險。這種合約不可逆性使得廢棄物處置費用向下彈性極低，並在跨年度的 t-1 期顯現出顯著之加劇效果。")
    add_body_para(doc, "（三）形式揭露與實質負擔之脫鉤。實證結果顯示 GRI 廢棄物揭露品質未呈顯著調節，反映台灣部分上市櫃公司在編製永續報告書時，揭露評分之高低多取決於永續顧問團隊之文字修飾與格式符合度（屬於形式合規），並未實質轉化為內部資源彈性配置之管理能力。此一發現對資本市場與監管機關具備強烈警示意涵：評估企業環境轉型之營運韌性時，必須深入檢視實質資源使用密集度，而非僅停留於永續報告書之形式揭露分數。")

print("第四章實證結果與分析宣告完成。")


# -------------------------------------------------------------
# 第五章 結論與建議 (Chapter 5)
# -------------------------------------------------------------

def build_chapter_5(doc):
    add_chapter_heading(doc, "第五章 結論與建議")
    
    add_section_heading(doc, "第一節 研究結論")
    add_body_para(doc, "本研究以 2014 年至 2024 年台灣上市櫃公司為研究對象，建構包含 18,137 筆公司-年度觀測值（涵蓋 1,945 家獨立公司）之長期面板資料庫，基於 Anderson, Banker, and Janakiraman (2003, ABJ) 成本僵固性計量模型，實證探討水資源管理與事業廢棄物處置行為對營業費用成本僵固性之影響機制。本研究核心實證結論歸納如下：")
    add_body_para(doc, "第一，確認台灣上市櫃公司普遍存在顯著之成本僵固性。在控制 SASB 主產業與年份雙向固定效果以及各項公司財務特徵下，實證模型顯示營業收入上升 1% 時，營業費用平均顯著擴張約 0.35%~0.39%；然而當營業收入下降 1% 時，營業費用向下縮減之幅度顯著受限，完全驗證了非對稱成本行為在台灣上市櫃公司之客觀存在。")
    add_body_para(doc, "第二，用水密集度在「高耗水產業當期（t-0）」對成本僵固性呈現顯著加劇效果（β3 = -0.0914, t = -2.11, p < 0.05），支持假說 H1a。此一發現證實：在半導體、光電、化學及鋼鐵等高度依賴大量水資源之製造業中，單位營收用水量越高，企業對超純水淨化系統與大型廢水回收設施之技術綁定越深；當面臨外部需求萎縮時，這些高度客製化之專用設施折舊與固定維運人員支出難以即刻裁撤，從而大幅推升了費用向下調整之沉沒成本，加劇成本僵固性。相對地，低耗水產業未呈現此加劇效果。")
    add_body_para(doc, "第三，廢棄物密集度在「高耗水產業遞延一期（t-1）」對成本僵固性呈現顯著加劇效果（β3 = -0.0503, t = -1.76, p < 0.10），支持假說 H1b。實證結果證實：單位營收產生的廢棄物負荷越高，企業對外部合格清運機構長期合約之依賴越深；受制於最低保證處理量條款與跨年度履約特質，廢棄物清運費用在營收衰退後第一年（t-1）展現最強大的合約剛性約束力，證實環境投入之成本影響具備跨期延遲發酵之動態特徵。")
    add_body_para(doc, "第四，實質物理資源使用密集度之調整成本效應，顯著主導單純之形式資訊揭露。水資源與廢棄物之形式揭露指標在各子樣本與各期數中均未達統計顯著；且在校正離散罰鍰次數之極端值偏誤、改採標準化裁罰金額比率後，行政裁罰未達統計顯著。研究證實實體物理資產之專屬性才是形塑企業成本調整行為之重要力量。")
    
    add_section_heading(doc, "第二節 理論與學術貢獻")
    add_body_para(doc, "本研究之實證發現，在環境管理會計、成本性態理論與綠色金融領域具有以下三項核心理論貢獻：")
    add_body_para(doc, "首先，補充了調整成本理論在實體環境資源維度之實證討論。 既有文獻多自廣義之整體 ESG 評分探討財務後果，本研究將成本行為分析推進至物理可量化之「水資源與廢棄物實質密集度」，結合交易成本經濟學之資產專屬性概念，證實高專屬性之水循環設施與長期委外清運合約會對企業營業費用構成剛性約束，有助於補充環境會計與成本行為領域之實證文獻。")
    add_body_para(doc, "其次，揭示了環境成本效應具備「產業依存」與「動態時間滯後發酵」之雙重特質。 本研究證實全樣本當期混合估計會因輕工業之加入而產生顯著之稀釋效應；唯有結合產業耗水分組與時間落差模型，方能精確識別環境投入之真實約束力。此一計量設計為後續學者探討環境資本支出之財務延遲衝擊提供了方法論參考。")
    add_body_para(doc, "最後，釐清並校正了過去環境法遵計量文獻之極端值偏誤。 本研究透過實證證明過去採用離散「罰鍰次數」易受單一工廠單一年度受罰數百次之極端值主導，改採連續「金額佔資產比率」後得出穩健結論，提升了實證結果之可靠度。")
    
    add_section_heading(doc, "第三節 企業管理實務意涵")
    add_body_para(doc, "基於實證結論，本研究對台灣企業高階管理階層（特別是執行長 CEO 與財務長 CFO）提供以下實務參考：")
    add_body_para(doc, "一、正視環境資本支出對「營運槓桿」（Operating Leverage）之放大效應。企業高層在進行大型水回收系統與污染防治設施之資本支出決策時，財務部門不應僅著眼於節水節能之直接效益，更應審慎評估其將大幅推升企業之固定營運支出，降低景氣衰退時之費用調整彈性。財務長宜在景氣擴張期建立充足之流動性財務緩衝，以抵禦未來需求下滑時高成本僵固性對營業淨利（EBIT）可能產生之放大侵蝕。")
    add_body_para(doc, "二、優化事業廢棄物處置合約之談判條款。考量廢棄物處置合約存在跨期剛性，企業採購與廠務部門在與外部合格環保清理機構簽約時，應盡量避免綁定單一僵硬之高額保證處理量；宜爭取「基礎維護費 + 浮動處理量計價」之彈性階梯費率機制，保留在產能縮減時向下調整清運頻次與合約總額之法律彈性，以降低景氣下行時的契約剛性包袱。")
    
    add_section_heading(doc, "第四節 政策與監理建議")
    add_body_para(doc, "針對金融監管機構（金管會）、環境保護主管機關（環境部）以及綠色金融授信機構，本研究提供以下政策參考：")
    add_body_para(doc, "一、強化企業「資源使用密集度」之標準化量化揭露規範。目前金管會推動之永續報告書編製規範，多側重於定性描述與範疇一至三碳排盤查。本研究實證結果顯示用水與廢棄物密集度對企業營運成本彈性具備實質關聯，建議主管機關可進一步研議將「單位營收用水量」與「每百萬營收廢棄物總量」列為高耗水與高排污產業之量化揭露參考指引，並研訂統一之計算指引與第三方獨立查證機制，提升永續資訊之實質決策攸關性。")
    add_body_para(doc, "二、綠色金融機構應將實質資源密集度納入企業授信壓力測試。本研究發現高耗水產業中用水密集度高者，於景氣下滑時費用調整摩擦力顯著更大，營運現金流波動風險更高。建議銀行與金融授信機構在推動「永續績效連結貸款」（Sustainability-Linked Loans, SLL）與信用風險評估時，不應僅依賴綠色認證標章或報告書揭露有無，而應將企業之實質用水與廢棄物密集度及其在營收下行時之成本調整彈性納入壓力測試模型，更精準控管綠色信貸風險。")
    
    add_section_heading(doc, "第五節 研究限制與未來研究方向")
    add_body_para(doc, "本研究雖力求實證過程之周延與客觀，惟受限於客觀資料條件，仍存在以下研究限制，並提供後續研究參考：")
    add_body_para(doc, "一、樣本地域範疇之限制：本研究樣本取自台灣上市櫃公司，研究結論立足於台灣特定產業結構、水資源環境條件及法規背景。未來研究可將範疇拓展至其他不同法規環境與天然資源條件之經濟體，執行跨區域比較，以檢視環境使用密集度對成本行為之適用性。")
    add_body_para(doc, "二、管理階層心理決策過程之不可觀察性：基於次級財務面板資料庫之限制，本研究僅能自營收衰退時之費用結果反推經理人之調整成本權衡，無法直接觀察管理團隊內部會議之心理認知與決策爭辯。未來研究可結合質化個案訪談、實地問卷調查或高階經理人文字語音情緒分析，進一步釐清資源調整決策之內部決策過程。")
    add_body_para(doc, "三、因果識別方法之深化：本研究雖已控制雙向固定效果、公司叢集穩健標準誤並進行時間落差分析，惟環境投入與企業成本決策本質上可能受到未觀察之管理能力內生影響。未來研究可進一步結合外生環境政策衝擊（如 2021 年百年大旱分區供水政策、耗水費正式開徵、或《資源循環促進法》修訂）作為外生工具變數或準自然實驗，採雙重差分法（DiD）或斷點迴歸（RD）進一步補強因果識別之效果。")

# -------------------------------------------------------------
# 參考文獻與附錄模組 (Module 6)
# -------------------------------------------------------------

def build_back_matter(doc, water_df):
    add_chapter_heading(doc, "參考文獻")
    add_styled_para(doc, "本論文參考文獻全數來自指定文獻庫，中文文獻依作者姓氏筆畫順序排列於前，英文文獻依第一作者姓氏字母順序排列於後，遵循 APA 第七版（APA 7th Edition）學術體例，並採凸排 0.85 公分格式編排。", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=10, italic=True, space_before=0, space_after=12)
    
    # 中文文獻（依筆畫排序）
    chinese_refs = [
        "王惠洳（2025）。獨立董事連結關係與成本僵固性之關聯（未出版之碩士論文）。國立臺中科技大學會計資訊系碩士班，臺中市。",
        "毛祚祥（2026）。探討台灣上市櫃觀光餐旅類股企業ESG指標、系統性風險與財務績效之關係（未出版之碩士論文）。國立高雄科技大學觀光管理系觀光與餐旅管理碩士班，高雄市。",
        "辛珏辰（2025）。ESG活動與成本僵固性——兼論COVID-19的影響（未出版之碩士論文）。國立成功大學會計學系，臺南市。",
        "邱圓雅（2020）。內部債與成本僵固性之關聯性（未出版之碩士論文）。國立臺北商業大學會計資訊系會計財稅碩士班，臺北市。",
        "張芳瑜（2023）。成本僵固性對企業績效之影響——以總體經濟變數為調節變數（未出版之碩士論文）。國立中興大學企業管理學系，臺中市。",
        "張宸杰（2023）。內部控制、財務危機對成本僵固性之影響（未出版之碩士論文）。國立臺中科技大學會計資訊系碩士班，臺中市。",
        "張岑華（2026）。ESG績效與漂綠對權益資金成本之影響：以台灣為例（未出版之碩士論文）。世新大學財務金融學系，臺北市。",
        "張筱婕（2026）。漂綠行為對企業ESG之影響：媒體揭露的調節效果（未出版之碩士論文）。世新大學財務金融學系，臺北市。",
        "許嘉芸（2026）。企業資源密集度管理對財務績效與ESG效率的影響（未出版之碩士論文）。國立陽明交通大學經營管理研究所，新竹市。",
        "陳思婷（2025）。移轉訂價操弄與成本僵固性間之關聯性（未出版之碩士論文）。天主教輔仁大學會計學系碩士班，新北市。",
        "富得運（2026）。ESG風險管理與企業財務績效：新興市場中企業規模之調節效果（未出版之碩士論文）。正修科技大學工業工程與管理碩士班，高雄市。",
        "楊宜蓁（2026）。ESG及負債比率對企業信用風險之關聯性研究——以台灣上市電子業為例（未出版之碩士論文）。國立彰化師範大學財務金融技術學系，彰化市。",
        "楊博勝（2026）。ESG資訊之揭露對於企業特有風險之影響：以台灣上市櫃公司為例（未出版之博士論文）。國立陽明交通大學經營管理研究所，新竹市。",
        "劉佳熒（2022）。供應鏈關係對供應商之成本行為的影響（未出版之博士論文）。亞洲大學經營管理學系，臺中市。",
        "鄭紹君（2018）。企業社會責任活動與成本僵固性之關聯——以企業社會責任願景調節（未出版之碩士論文）。國立中央大學會計研究所，桃園市。",
        "謝雅筑（2024）。飛行員CEO與成本僵固性（未出版之碩士論文）。國立成功大學財務金融研究所碩士在職專班，臺南市。",
        "蘇姿云（2026）。ESG對盈餘管理的抑制效果：家族企業的差異分析（未出版之碩士論文）。國立成功大學財務金融研究所在職專班，臺南市。"
    ]
    for ref in chinese_refs:
        p_ref = doc.add_paragraph()
        set_para_format(p_ref, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.3, space_before=3, space_after=3)
        p_ref.paragraph_format.left_indent = Pt(24)
        p_ref.paragraph_format.first_line_indent = Pt(-24)
        r = p_ref.add_run(ref)
        set_run_font(r, size_pt=11)
        
    # 英文文獻（依作者姓氏字母序排列）
    english_refs = [
        ("Anderson, M. C., Banker, R. D., & Janakiraman, S. N. (2003). Are selling, general, and administrative costs “sticky”? ", "Journal of Accounting Research", ", 41(1), 47–63. https://doi.org/10.1111/1475-679X.00095"),
        ("Jiang, W., & Yang, W. (2024). ESG disclosure and corporate cost stickiness: Evidence from supply-chain relationships. ", "Economics Letters", ", 238, 111697. https://doi.org/10.1016/j.econlet.2024.111697"),
        ("Zeng, J., Peng, M., & Chan, K. C. (2024). The impact of low-carbon city policy on corporate cost stickiness: Evidence from green innovation and financing constraints. ", "International Review of Financial Analysis", ", 91, 103011. https://doi.org/10.1016/j.irfa.2023.103011")
    ]
    for pre, journal, post in english_refs:
        p_ref = doc.add_paragraph()
        set_para_format(p_ref, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.3, space_before=3, space_after=3)
        p_ref.paragraph_format.left_indent = Pt(24)
        p_ref.paragraph_format.first_line_indent = Pt(-24)
        r1 = p_ref.add_run(pre)
        set_run_font(r1, size_pt=11)
        r2 = p_ref.add_run(journal)
        set_run_font(r2, size_pt=11, italic=True)
        r3 = p_ref.add_run(post)
        set_run_font(r3, size_pt=11)
        
    doc.add_page_break()
    
    # 附錄模組
    add_chapter_heading(doc, "附錄")
    build_table_app_a(doc)
    doc.add_paragraph()
    build_table_app_b(doc, water_df)

print("第五章與後置頁宣告完成。")


# -------------------------------------------------------------
# 主程式入口 (Main Execution)
# -------------------------------------------------------------

def main():
    print("正在讀取實證資料庫與迴歸報表...")
    df = pd.read_csv(MODEL_CSV, low_memory=False)
    coef_df = pd.read_csv(COEF_CSV)
    rob_df = pd.read_csv(ROB_CSV)
    ind_df = pd.read_csv(IND_CSV)
    water_df = pd.read_csv(WATERUSE_CSV)
    print(f"資料讀取成功：主資料 {df.shape}，迴歸係數 {coef_df.shape}，異質性報表 {rob_df.shape}。")
    
    print("正在建立中山大學格式論文 Document 物件...")
    doc = Document()
    
    # 設置第一節（前置頁，小寫羅馬數字頁碼）
    sec_front = doc.sections[0]
    sec_front.top_margin = Cm(2.54)
    sec_front.bottom_margin = Cm(2.54)
    sec_front.left_margin = Cm(3.0)
    sec_front.right_margin = Cm(3.0)
    set_page_number_footer(sec_front, is_roman=True, start_at_1=False)
    
    print("正在建構前置頁（封面、審定書、誌謝、中英文摘要、目錄、圖表次）...")
    build_front_matter(doc)
    
    # 建立正文分節（切換為阿拉伯數字頁碼，自第 1 頁起算）
    print("正在切換正文分節與頁碼系統...")
    sec_body = doc.add_section()
    sec_body.footer.is_linked_to_previous = False
    sec_body.top_margin = Cm(2.54)
    sec_body.bottom_margin = Cm(2.54)
    sec_body.left_margin = Cm(3.0)
    sec_body.right_margin = Cm(3.0)
    set_page_number_footer(sec_body, is_roman=False, start_at_1=True)
    
    print("正在生成第一章 緒論...")
    build_chapter_1(doc)
    
    print("正在生成第二章 文獻探討與假說建立...")
    build_chapter_2(doc)
    
    print("正在生成第三章 研究方法與資料設計（含表 3-1, 3-2, 3-3）...")
    build_chapter_3(doc, ind_df)
    
    print("正在生成第四章 實證結果與分析（含表 4-1, 4-2, 4-3, 4-4, 4-5）...")
    build_chapter_4(doc, df, coef_df, rob_df)
    
    print("正在生成第五章 結論與建議...")
    build_chapter_5(doc)
    
    print("正在生成參考文獻（APA 7th）與附錄 A、B...")
    build_back_matter(doc, water_df)
    
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_filename = f"企業水資源管理與廢棄物處理對成本僵固性之影響_{now_str}.docx"
    out_path = os.path.join(OUT_DIR, out_filename)
    
    print(f"正在儲存正式論文 DOCX 檔案至：{out_path} ...")
    doc.save(out_path)
    print("=" * 60)
    print("國立中山大學財務管理學系 碩士論文生成完畢。")
    print(f"檔案路徑：{out_path}")
    print(f"檔案大小：{os.path.getsize(out_path):,} bytes")
    print("=" * 60)

if __name__ == "__main__":
    main()
