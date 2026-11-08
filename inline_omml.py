# -*- coding: utf-8 -*-
"""
inline_omml.py
--------------
Word 原生行內公式 (Inline OMML) 處理模組。
將論文內文、標題、附註與表格中的計量經濟學變數與數學符號，
轉換為 Word 原生 <m:oMath> 行內公式物件，確保下標、符號與字型格式 100% 精確一致。
"""

import re
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor

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

# 核心變數之精確 OMML 定義 (Cambria Math 字型標準)
EXACT_MAP = {
    # 應變數與自變數
    'ΔlnSGA_{i,t}': '<m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:sSub><m:e><m:r><m:t>SGA</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'ΔlnSGA': '<m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>SGA</m:t></m:r>',
    'ΔlnREV_{i,t}': '<m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'ΔlnREV': '<m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    
    # 含有 ln 之定義式
    'ln(SGA_{i,t})': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>SGA</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'ln(SGA_{i,t-1})': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>SGA</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-1</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'ln(REV_{i,t})': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'ln(REV_{i,t-1})': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-1</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'ln(Total Assets)': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>(Total Assets)</m:t></m:r>',
    
    # 水準變數與落後變數
    'SGA_{i,t}': '<m:sSub><m:e><m:r><m:t>SGA</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'SGA_{i,t-1}': '<m:sSub><m:e><m:r><m:t>SGA</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-1</m:t></m:r></m:sub></m:sSub>',
    'REV_{i,t}': '<m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'REV_{i,t-1}': '<m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-1</m:t></m:r></m:sub></m:sSub>',
    'DEC_{i,t}': '<m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'DEC': '<m:r><m:t>DEC</m:t></m:r>',
    
    # 環境調節變數
    'Env_Var_{i,t}': '<m:sSub><m:e><m:r><m:t>Env_Var</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'Env_Var_{i,t-k}': '<m:sSub><m:e><m:r><m:t>Env_Var</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-k</m:t></m:r></m:sub></m:sSub>',
    'Env_Var_{i,t-m}': '<m:sSub><m:e><m:r><m:t>Env_Var</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-m</m:t></m:r></m:sub></m:sSub>',
    'Env_Var': '<m:r><m:t>Env_Var</m:t></m:r>',
    'WATER_INTENSITY_{i,t}': '<m:sSub><m:e><m:r><m:t>WATER_INTENSITY</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'WASTE_INTENSITY_{i,t}': '<m:sSub><m:e><m:r><m:t>WASTE_INTENSITY</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'WATER_INTENSITY': '<m:r><m:t>WATER_INTENSITY</m:t></m:r>',
    'WASTE_INTENSITY': '<m:r><m:t>WASTE_INTENSITY</m:t></m:r>',
    'WATER_RATE': '<m:r><m:t>WATER_RATE</m:t></m:r>',
    'WATER_DISC': '<m:r><m:t>WATER_DISC</m:t></m:r>',
    'WASTE_DISC': '<m:r><m:t>WASTE_DISC</m:t></m:r>',
    'WASTE_FINE': '<m:r><m:t>WASTE_FINE</m:t></m:r>',
    'HIGH_WATER': '<m:r><m:t>HIGH_WATER</m:t></m:r>',
    
    # 控制變數
    'SUCC_DEC': '<m:r><m:t>SUCC_DEC</m:t></m:r>',
    'SIZE': '<m:r><m:t>SIZE</m:t></m:r>',
    'AI': '<m:r><m:t>AI</m:t></m:r>',
    'EI': '<m:r><m:t>EI</m:t></m:r>',
    'ROA': '<m:r><m:t>ROA</m:t></m:r>',
    'LEV': '<m:r><m:t>LEV</m:t></m:r>',
    
    # 迴歸係數與固定效果
    'β0': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>',
    'β1': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>',
    'β2': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub>',
    'β3': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>3</m:t></m:r></m:sub></m:sSub>',
    'β4': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>4</m:t></m:r></m:sub></m:sSub>',
    'β_{3,lag}': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>3,lag</m:t></m:r></m:sub></m:sSub>',
    'β_{4,lag}': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>4,lag</m:t></m:r></m:sub></m:sSub>',
    'α_{ind}': '<m:sSub><m:e><m:r><m:t>α</m:t></m:r></m:e><m:sub><m:r><m:t>ind</m:t></m:r></m:sub></m:sSub>',
    'γ_{yr}': '<m:sSub><m:e><m:r><m:t>γ</m:t></m:r></m:e><m:sub><m:r><m:t>yr</m:t></m:r></m:sub></m:sSub>',
    'ε_{i,t}': '<m:sSub><m:e><m:r><m:t>ε</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    
    # 典型交乘項
    'DEC_{i,t} × ΔlnREV_{i,t}': '<m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    '(DEC_{i,t} × ΔlnREV_{i,t})': '<m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'DEC_{i,t} × ΔlnREV': '<m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'Env × DEC_{i,t} × ΔlnREV_{i,t}': '<m:r><m:t>Env</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub>',
    'Env × DEC_{i,t} × ΔlnREV': '<m:r><m:t>Env</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'SIZE × DEC_{i,t} × ΔlnREV': '<m:r><m:t>SIZE</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'AI × DEC_{i,t} × ΔlnREV': '<m:r><m:t>AI</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'EI × DEC_{i,t} × ΔlnREV': '<m:r><m:t>EI</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'ROA × DEC_{i,t} × ΔlnREV': '<m:r><m:t>ROA</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'LEV × DEC_{i,t} × ΔlnREV': '<m:r><m:t>LEV</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    'SUCC_DEC × DEC_{i,t} × ΔlnREV': '<m:r><m:t>SUCC_DEC</m:t></m:r><m:r><m:t>×</m:t></m:r><m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>×</m:t></m:r><m:r><m:t>Δ</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>ln</m:t></m:r><m:r><m:t>REV</m:t></m:r>',
    
    # 常用關係與統計量
    'REV_{i,t} < REV_{i,t-1}': '<m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>&lt;</m:t></m:r><m:sSub><m:e><m:r><m:t>REV</m:t></m:r></m:e><m:sub><m:r><m:t>i,t-1</m:t></m:r></m:sub></m:sSub>',
    'DEC_{i,t} = 1': '<m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>=1</m:t></m:r>',
    'DEC_{i,t} = 0': '<m:sSub><m:e><m:r><m:t>DEC</m:t></m:r></m:e><m:sub><m:r><m:t>i,t</m:t></m:r></m:sub></m:sSub><m:r><m:t>=0</m:t></m:r>',
    'β1 > 0': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub><m:r><m:t>&gt;0</m:t></m:r>',
    'β2 < 0': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub><m:r><m:t>&lt;0</m:t></m:r>',
    'β3 < 0': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>3</m:t></m:r></m:sub></m:sSub><m:r><m:t>&lt;0</m:t></m:r>',
    'β3 > 0': '<m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>3</m:t></m:r></m:sub></m:sSub><m:r><m:t>&gt;0</m:t></m:r>',
    '(β1 + β2)': '<m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub><m:r><m:t>+</m:t></m:r><m:sSub><m:e><m:r><m:t>β</m:t></m:r></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>',
    'R²': '<m:sSup><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    'R2': '<m:sSup><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    'Adj. R²': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>Adj. </m:t></m:r><m:sSup><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    'Adj. R2': '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>Adj. </m:t></m:r><m:sSup><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
}

def math_to_omath_xml(expr):
    expr = expr.strip()
    if expr in EXACT_MAP:
        body = EXACT_MAP[expr]
    else:
        # 通用簡易處理：解析下標 _ 與上標 ^
        if '_' in expr:
            parts = expr.split('_', 1)
            main_p = parts[0]
            sub_p = parts[1].strip('{}')
            body = f'<m:sSub><m:e><m:r><m:t>{main_p}</m:t></m:r></m:e><m:sub><m:r><m:t>{sub_p}</m:t></m:r></m:sub></m:sSub>'
        elif '^' in expr:
            parts = expr.split('^', 1)
            main_p = parts[0]
            sup_p = parts[1].strip('{}')
            body = f'<m:sSup><m:e><m:r><m:t>{main_p}</m:t></m:r></m:e><m:sup><m:r><m:t>{sup_p}</m:t></m:r></m:sup></m:sSup>'
        else:
            body = f'<m:r><m:t>{expr}</m:t></m:r>'
    return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{body}</m:oMath>'

# 依長度遞減排序以避免子字串提前比對
AUTO_MATH_TOKENS = sorted(list(EXACT_MAP.keys()), key=len, reverse=True)

# 針對純英文字母單字使用 word boundary，其他使用 re.escape
patterns = []
for tok in AUTO_MATH_TOKENS:
    if tok.isalpha():
        patterns.append(r'\b' + re.escape(tok) + r'\b')
    else:
        patterns.append(re.escape(tok))

MATH_PATTERN = re.compile(r'(\$(?:[^\$]+)\$|' + '|'.join(patterns) + r')')

def append_text_and_math(p, text, font_name_cjk=CJK_FONT, font_name_en=EN_FONT, size_pt=12, bold=False, italic=False, color_rgb=(0,0,0)):
    """
    將文字拆解為一般文字 Run 與 Word 原生 <m:oMath> 行內公式物件，依序加入段落。
    """
    if text is None:
        return
    text = str(text)
    if not text:
        return
    
    tokens = MATH_PATTERN.split(text)
    for token in tokens:
        if not token:
            continue
        if token.startswith('$') and token.endswith('$'):
            expr = token[1:-1]
            xml = math_to_omath_xml(expr)
            p._p.append(parse_xml(xml))
        elif token in EXACT_MAP:
            xml = math_to_omath_xml(token)
            p._p.append(parse_xml(xml))
        else:
            run = p.add_run(token)
            set_run_font(run, font_name_cjk=font_name_cjk, font_name_en=font_name_en, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
