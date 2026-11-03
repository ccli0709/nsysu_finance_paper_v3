import os
import re

input_file = r'd:\Repositories\nsysu_finance_paper_v3\08-中山格式論文生成\08_generate_thesis_template.py'
with open(input_file, 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Imports
code = code.replace('from docx import Document\nfrom docx.shared import Pt\nfrom docx.enum.text import WD_ALIGN_PARAGRAPH\nfrom docx.oxml.ns import qn\nfrom docx.oxml import OxmlElement\n', '')

# 2. Output file
code = code.replace('OUT = os.path.join(BASE_DIR, "08-中山格式論文生成", "水與廢棄物管理與成本黏性_論文_中山格式.docx")', 
                    'OUT_MD = os.path.join(BASE_DIR, "08-中山格式論文生成", "水與廢棄物管理與成本黏性_論文_中山格式.md")')

# 3. docx helpers -> markdown helpers
helpers = """
def font(run, size=12, bold=False):
    pass

def H1(md_lines, text):
    md_lines.append("# " + text + "\\\\n\\\\n")

def H2(md_lines, text):
    md_lines.append("## " + text + "\\\\n\\\\n")

def H3(md_lines, text):
    md_lines.append("### " + text + "\\\\n\\\\n")

def body(md_lines, text, align=None, indent=True, bold=False):
    if bold:
        text = f"**{text}**"
    md_lines.append(text + "\\\\n\\\\n")

def center(md_lines, text, size=16, bold=False):
    if bold:
        text = f"**{text}**"
    md_lines.append("<center>" + text + "</center>\\\\n\\\\n")

def caption(md_lines, text):
    md_lines.append(f"**{text}**\\\\n\\\\n")

def _set_border(cell, edge, sz=6):
    pass

def three_line(tbl, header_rows=1):
    pass

def fill(cell, text, size=10, bold=False, align="center"):
    pass
"""

code = re.sub(r'# ---------------- 字型/樣式工具 ----------------.*?# ---------------- 讀資料 ----------------', 
              '# ---------------- 字型/樣式工具 ----------------\n' + helpers + '\n# ---------------- 讀資料 ----------------', 
              code, flags=re.DOTALL)

# 4. Table Builders (Rewrite to generate Markdown tables directly)
table_builders = """
def desc_table(md_lines):
    rows = [["變數", "N", "Mean", "SD", "Min", "P25", "Median", "P75", "Max"]]
    for col, lab in DESC:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        rows.append([lab, f"{s.size:,}", f"{s.mean():.3f}", f"{s.std():.3f}",
                     f"{s.min():.3f}", f"{s.quantile(.25):.3f}", f"{s.median():.3f}",
                     f"{s.quantile(.75):.3f}", f"{s.max():.3f}"])
    md_lines.append("| " + " | ".join(rows[0]) + " |\\\\n")
    md_lines.append("|" + "|".join(["---"] * len(rows[0])) + "|\\\\n")
    for r in rows[1:]:
        md_lines.append("| " + " | ".join(str(x).replace('\\\\n',' ') for x in r) + " |\\\\n")
    md_lines.append("\\\\n")

def corr_table(md_lines):
    cols = [c for c, _ in CORR]
    labs = [l for _, l in CORR]
    sub = df[cols].apply(pd.to_numeric, errors="coerce")
    cm = sub.corr()
    n = len(sub.dropna())
    head = ["變數"] + [f"({k+1})" for k in range(len(cols))]
    md_lines.append("| " + " | ".join(head) + " |\\\\n")
    md_lines.append("|" + "|".join(["---"] * len(head)) + "|\\\\n")
    for i, lab in enumerate(labs):
        row = [f"({i+1}) {lab}"]
        for j in range(len(cols)):
            if j > i:
                row.append("")
                continue
            r = cm.iloc[i, j]
            if i == j:
                row.append("1.000")
            else:
                t = abs(r) * np.sqrt(max(n - 2, 1) / max(1e-9, 1 - r * r))
                val = f"{r:.3f}"
                if t >= 1.96:
                    val = f"**{val}**"
                row.append(val)
        md_lines.append("| " + " | ".join(row) + " |\\\\n")
    md_lines.append("\\\\n")

def reg_main_table(md_lines):
    ncol = 1 + len(L0_MODELS)
    head = ["變數"] + [lab.replace('\\\\n', ' ') for _, lab in L0_MODELS]
    md_lines.append("| " + " | ".join(head) + " |\\\\n")
    md_lines.append("|" + "|".join(["---"] * len(head)) + "|\\\\n")
    
    for greek, label, var in ROLE_ROWS:
        role = {"β1": "β1", "β2": "β2(黏性)", "β3": "β3(三重交乘)", "β4": "β4(水主效果)"}.get(greek)
        disp = f"{label}" + (f" ({greek})" if greek else "")
        
        row_b = [disp]
        row_t = [""]
        for m, _ in L0_MODELS:
            b, t = _coef_cell(m, role=role, var=var)
            row_b.append(b)
            row_t.append(t)
        md_lines.append("| " + " | ".join(row_b) + " |\\\\n")
        md_lines.append("| " + " | ".join(row_t) + " |\\\\n")
    
    row_n = ["N"]
    for m, _ in L0_MODELS:
        r = coef[coef["模型"] == m]
        row_n.append(f"{int(r['N'].iloc[0]):,}" if len(r) else "")
    md_lines.append("| " + " | ".join(row_n) + " |\\\\n")
    
    row_r2 = ["adj. R²"]
    for m, _ in L0_MODELS:
        r = coef[coef["模型"] == m]
        row_r2.append(f"{r['adj_R2'].iloc[0]:.4f}" if len(r) else "")
    md_lines.append("| " + " | ".join(row_r2) + " |\\\\n")
    md_lines.append("\\\\n")

def lag_beta3_table(md_lines):
    b3 = coef[coef["角色"] == "β3(三重交乘)"]
    kinds = [("用水密集度", "模型6 用水密集度"),
             ("廢棄物主分析", "模型3 廢棄物密集度"),
             ("主分析", "模型1 水回收率%"), ("次分析", "模型2 水揭露"),
             ("廢棄物次分析", "模型4 廢棄物揭露"), ("廢棄物穩健", "模型5 廢棄物裁罰金額")]
    lags = sorted(b3["遞延期"].unique())
    head = ["遞延期"] + [lab for _, lab in kinds]
    md_lines.append("| " + " | ".join(head) + " |\\\\n")
    md_lines.append("|" + "|".join(["---"] * len(head)) + "|\\\\n")
    for lg in lags:
        row = [f"t-{lg}"]
        for kind, _ in kinds:
            r = b3[(b3["遞延期"] == lg) & (b3["類型"] == kind)]
            if r.empty:
                row.append("")
            else:
                b = r["係數"].iloc[0]; s = r["顯著性"].iloc[0]
                s = "" if pd.isna(s) else str(s)
                row.append(f"{b:.4f}{s}")
        md_lines.append("| " + " | ".join(row) + " |\\\\n")
    md_lines.append("\\\\n")

def hetero_table(md_lines):
    cols = ["衡量", "產業組", "遞延期", "β2", "β3", "N"]
    md_lines.append("| " + " | ".join(cols) + " |\\\\n")
    md_lines.append("|" + "|".join(["---"] * len(cols)) + "|\\\\n")
    for _, r in rob.iterrows():
        b2 = "" if pd.isna(r["β2(D×ΔREV)"]) else f"{r['β2(D×ΔREV)']:.4f}{'' if r['β2_sig']=='n.s.' else r['β2_sig']}"
        b3 = "" if pd.isna(r["β3(Water×D×ΔREV)"]) else f"{r['β3(Water×D×ΔREV)']:.4f}{'' if r['β3_sig'] in ('n.s.','樣本不足') else r['β3_sig']}"
        vals = [r["水衡量"], r["產業組"], str(r["遞延期"]), b2, b3, f"{int(r['N']):,}"]
        md_lines.append("| " + " | ".join(vals) + " |\\\\n")
    md_lines.append("\\\\n")

def var_def_table(md_lines):
    rows = [
        ("ΔLNSGA", "LN(本期營業費用) − LN(前期營業費用)"),
        ("ΔLNREV", "LN(本期營業收入淨額) − LN(前期營業收入淨額)"),
        ("DEC", "收入下降虛擬：本期營收<前期營收=1，否則0"),
        ("WATER_RATE", "水回收率%（TESG 企業用水量）"),
        ("WATER_DISC", "水揭露虛擬：當年有水資料紀錄=1"),
        ("WASTE_INTENSITY", "每百萬營收廢棄物量（TESG 廢棄物揭露）"),
        ("WASTE_DISC", "GRI 廢棄物管理揭露度（揭露品質）"),
        ("WASTE_FINE", "事業廢棄物罰鍰次數"),
        ("SIZE", "LN(資產總額)"),
        ("AI", "資產密集度＝資產總額/營業收入淨額"),
        ("EI", "員工密集度＝員工人數/營業收入淨額"),
        ("ROA", "資產報酬率 ROA(A)稅後息前"),
        ("LEV", "財務槓桿＝負債比率"),
        ("SUCC_DEC", "連續兩年營收下降虛擬"),
    ]
    md_lines.append("| 變數名稱 | 定義 |\\\\n")
    md_lines.append("|---|---|\\\\n")
    for a, b in rows:
        md_lines.append(f"| {a} | {b} |\\\\n")
    md_lines.append("\\\\n")
"""

code = re.sub(r'# ---------------- 表格建構 ----------------.*?# ---------------- 產生文件 ----------------', 
              '# ---------------- 表格建構 ----------------\n' + table_builders + '\n# ---------------- 產生文件 ----------------', 
              code, flags=re.DOTALL)

# 5. In main(): doc = Document() -> md_lines = []
code = re.sub(r'    doc = Document\(\)\n.*?n\.element\.rPr\.rFonts\.set\(qn\("w:eastAsia"\), CJK\)\n', '    md_lines = []\n', code, flags=re.DOTALL)

# Change doc.add_paragraph() and doc.add_page_break() to pass
code = code.replace('doc.add_paragraph()', '')
code = code.replace('doc.add_page_break()', '')

# Change all `doc` to `md_lines`
code = code.replace('H1(doc,', 'H1(md_lines,')
code = code.replace('H2(doc,', 'H2(md_lines,')
code = code.replace('H3(doc,', 'H3(md_lines,')
code = code.replace('body(doc,', 'body(md_lines,')
code = code.replace('center(doc,', 'center(md_lines,')
code = code.replace('caption(doc,', 'caption(md_lines,')
code = code.replace('desc_table(doc)', 'desc_table(md_lines)')
code = code.replace('corr_table(doc)', 'corr_table(md_lines)')
code = code.replace('reg_main_table(doc)', 'reg_main_table(md_lines)')
code = code.replace('lag_beta3_table(doc)', 'lag_beta3_table(md_lines)')
code = code.replace('hetero_table(doc)', 'hetero_table(md_lines)')
code = code.replace('var_def_table(doc)', 'var_def_table(md_lines)')

# 8. References
ref_code = """
    def add_ref(text):
        md_lines.append("- " + text + "\\\\n")
"""
code = re.sub(r'    def add_ref\(text\):.*?font\(r\)\n', ref_code, code, flags=re.DOTALL)

# 9. Save
code = code.replace('doc.save(OUT_MD)', 'with open(OUT_MD, "w", encoding="utf-8") as f:\\n        f.write("".join(md_lines))')
code = code.replace('OUT', 'OUT_MD')
code = code.replace('水與廢棄物管理與成本黏性_論文_中山格式.docx', '水與廢棄物管理與成本黏性_論文_中山格式.md')

# Restore globals deleted by regex
globals_code = """
ROLE_ROWS = [
    ("β1", "ΔLNREV", "dLNREV"),
    ("β2", "D×ΔLNREV", "DEC_x_dREV"),
    ("β3", "環境變數×D×ΔLNREV", None),
    ("β4", "環境變數", None),
    ("", "SIZE×D×ΔLNREV", "SIZE_D_dREV"),
    ("", "AI×D×ΔLNREV", "AI_D_dREV"),
    ("", "EI×D×ΔLNREV", "EI_D_dREV"),
    ("", "ROA×D×ΔLNREV", "ROA_D_dREV"),
    ("", "LEV×D×ΔLNREV", "LEV_D_dREV"),
    ("", "SUCC_DEC×D×ΔLNREV", "SUCC_DEC_D_dREV"),
]
def _coef_cell(model, role=None, var=None):
    r = coef[coef["模型"] == model]
    if role:
        rr = r[r["角色"] == role]
    else:
        rr = r[r["變數"] == var]
    if rr.empty:
        return "", ""
    b = rr["係數"].iloc[0]; s = rr["顯著性"].iloc[0]; t = rr["t值"].iloc[0]
    import pandas as pd
    s = "" if pd.isna(s) else str(s)
    return f"{b:.4f}{s}", f"({t:.2f})"
"""
code = code.replace('# ---------------- 表格建構 ----------------', '# ---------------- 表格建構 ----------------\\n' + globals_code)

with open(r'd:\Repositories\nsysu_finance_paper_v3\08-中山格式論文生成\08_generate_thesis_template.py', 'w', encoding='utf-8') as f:
    f.write(code)

