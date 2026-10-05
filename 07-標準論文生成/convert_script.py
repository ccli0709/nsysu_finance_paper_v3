import os

input_file = r'd:\Repositories\nsysu_finance_paper_v3\07-標準論文生成\07_generate_paper.py'
with open(input_file, 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Imports
code = code.replace('from docx import Document\nfrom docx.shared import Pt, RGBColor\nfrom docx.enum.text import WD_ALIGN_PARAGRAPH\nfrom docx.oxml.ns import qn\n', '')

# 2. Output file
code = code.replace('OUT_DOCX = os.path.join(BASE_DIR, "07-標準論文生成", "水與廢棄物管理與成本黏性_論文.docx")', 
                    'OUT_MD = os.path.join(BASE_DIR, "07-標準論文生成", "水與廢棄物管理與成本黏性_論文.md")')

# 3. docx helpers -> markdown helpers
helpers = """
def set_cjk(run):
    pass

def add_heading(md_lines, text, level=1):
    md_lines.append("#" * level + " " + text + "\\\\n")
    return None

def add_para(md_lines, text, size=12, bold=False, align=None, first_indent=True):
    if bold:
        text = f"**{text}**"
    md_lines.append(text + "\\\\n\\\\n")
    return None

def add_table(md_lines, df, title=None, num_fmt="{:.4f}"):
    if title:
        md_lines.append(f"**{title}**\\\\n")
    
    # Headers
    cols = df.columns.tolist()
    md_lines.append("| " + " | ".join(str(c) for c in cols) + " |")
    md_lines.append("|" + "|".join(["---"] * len(cols)) + "|")
    
    # Rows
    for _, row in df.iterrows():
        row_str = []
        for v in row:
            if isinstance(v, int):
                row_str.append(f"{v:,}")
            elif isinstance(v, float):
                import pandas as pd
                row_str.append("" if pd.isna(v) else num_fmt.format(v))
            else:
                row_str.append(str(v))
        md_lines.append("| " + " | ".join(row_str) + " |")
    md_lines.append("\\\\n")
    return None
"""

import re
code = re.sub(r'def set_cjk\(run\):.*?return t\n', helpers, code, flags=re.DOTALL)

# 4. In main(): doc = Document() -> md_lines = []
code = re.sub(r'doc = Document\(\)\n.*?normal\.element\.rPr\.rFonts\.set\(qn\("w:eastAsia"\), CJK_FONT\)\n', 'md_lines = []\n', code, flags=re.DOTALL)

# 5. doc.add_paragraph() stuff for title
title_code = """
    # 封面
    md_lines.append("# 國立中山大學財務管理學系碩士論文\\\\n## Master’s Thesis\\\\n## Department of Finance\\\\n## National Sun Yat-sen University\\\\n\\\\n# 企業水資源管理與廢棄物處理對成本黏性之影響\\\\n## Water Resource Management, Waste Treatment, and Corporate Cost Stickiness\\\\n")
    md_lines.append("**研究生：研究團隊**\\\\n**指導教授：封之遠 博士**\\\\n**中華民國115年6月**\\\\n**June 2026**\\\\n\\\\n")
"""
code = re.sub(r'    # 封面\n.*?doc\.add_paragraph\(\)\n', title_code, code, flags=re.DOTALL)

# 6. Change all `doc` to `md_lines`
code = code.replace('add_heading(doc,', 'add_heading(md_lines,')
code = code.replace('add_para(doc,', 'add_para(md_lines,')
code = code.replace('add_table(doc,', 'add_table(md_lines,')

# 7. add_ref logic
ref_code = """
    def add_ref(text):
        md_lines.append("- " + text + "\\\\n")

"""
code = re.sub(r'    def add_ref\(text\):.*?set_cjk\(r\)\n', ref_code, code, flags=re.DOTALL)

# 8. doc.save -> save to md
code = code.replace('doc.save(OUT_DOCX)', 'with open(OUT_MD, "w", encoding="utf-8") as f:\\n        f.write("\\\\n".join(md_lines))')
code = code.replace('OUT_DOCX', 'OUT_MD')

# 9. Clean up any leftover doc.add_paragraph()
code = code.replace('doc.add_paragraph()', '')

with open(r'd:\Repositories\nsysu_finance_paper_v3\07-標準論文生成\07_generate_paper_md.py', 'w', encoding='utf-8') as f:
    f.write(code)

