# -*- coding: utf-8 -*-
"""
08_generate_thesis_template.py
------------------------------
以 template_paper/氣候風險與企業碳排放_更01.docx 之架構、段落與格式為樣版，
把本研究（水＋廢棄物管理對成本黏性）的結果與數據製成新論文 Word。

格式重點（比照樣版）：
- 標楷體；本文 12pt；H1 置中 16pt 粗體、H2 14pt 粗體、H3 14pt。
- 章名不加「第X章」（緒論／文獻回顧與研究假說／資料描述與研究方法／研究結果／結論）。
- 迴歸表為三線表：變數列為係數＋顯著星號，下一列為括號內 t 值；表尾列 N、adj. R²。
- 敘述統計：Mean/SD/Min/P25/Median/P75/Max；相關矩陣：下三角、顯著者粗體。

輸出：水與廢棄物管理與成本黏性_論文_中山格式.docx
"""
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

MODEL_CSV = "04_model_data.csv"
COEF_CSV = "05_regression_coef.csv"
ROB_CSV = "06_robustness_summary.csv"
OUT = "水與廢棄物管理與成本黏性_論文_中山格式.docx"
CJK, EN = "標楷體", "Times New Roman"

# 主迴歸五模型（L0）欄位順序與標題
L0_MODELS = [
    ("模型1 水回收率% L0(當期)", "模型1\n水回收率%"),
    ("模型2 水揭露 L0(當期)", "模型2\n水揭露"),
    ("模型3 廢棄物密集度 L0(當期)", "模型3\n廢棄物密集度"),
    ("模型4 廢棄物揭露 L0(當期)", "模型4\n廢棄物揭露"),
    ("模型5 廢棄物罰鍰 L0(當期)", "模型5\n廢棄物罰鍰"),
]
DESC = [("Y_dLNSGA", "ΔLNSGA"), ("dLNREV", "ΔLNREV"), ("D", "D"),
        ("Water_Rate", "水回收率%"), ("Waste_Intensity", "每百萬營收廢棄物"),
        ("Waste_Disc", "廢棄物揭露度"), ("Waste_Fine", "廢棄物罰鍰次數"),
        ("Size", "Size"), ("AI", "AI"), ("EI", "EI"), ("ROA", "ROA"),
        ("Lev", "Lev"), ("Decrease", "Decrease")]
CORR = [("Y_dLNSGA", "ΔLNSGA"), ("dLNREV", "ΔLNREV"), ("Water_Rate", "水回收率%"),
        ("Waste_Intensity", "廢棄物密集度"), ("Waste_Disc", "廢棄物揭露"),
        ("Size", "Size"), ("AI", "AI"), ("ROA", "ROA"), ("Lev", "Lev")]


# ---------------- 字型/樣式工具 ----------------
def font(run, size=12, bold=False):
    run.font.name = EN
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)


def H1(doc, text):
    p = doc.add_heading(level=1); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run(text), 16, True); return p


def H2(doc, text):
    p = doc.add_heading(level=2); font(p.add_run(text), 14, True); return p


def H3(doc, text):
    p = doc.add_heading(level=3); font(p.add_run(text), 14, False); return p


def body(doc, text, align=None, indent=True, bold=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing = 1.5
    font(p.add_run(text), 12, bold); return p


def center(doc, text, size=16, bold=False):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run(text), size, bold); return p


def caption(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text); font(r, 12, True); return p


def _set_border(cell, edge, sz=6):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders"); tcPr.append(borders)
    e = borders.find(qn(f"w:{edge}"))
    if e is None:
        e = OxmlElement(f"w:{edge}"); borders.append(e)
    e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
    e.set(qn("w:space"), "0"); e.set(qn("w:color"), "000000")


def three_line(tbl, header_rows=1):
    """三線表：表頂線、表頭下線、表底線。"""
    nr = len(tbl.rows)
    for c in tbl.rows[0].cells:
        _set_border(c, "top", 12)
    for c in tbl.rows[header_rows - 1].cells:
        _set_border(c, "bottom", 6)
    for c in tbl.rows[nr - 1].cells:
        _set_border(c, "bottom", 12)


def fill(cell, text, size=10, bold=False, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    for i, line in enumerate(str(text).split("\n")):
        if i > 0:
            p.add_run().add_break()
        font(p.add_run(line), size, bold)


# ---------------- 讀資料 ----------------
df = pd.read_csv(MODEL_CSV, encoding="utf-8-sig", low_memory=False)
coef = pd.read_csv(COEF_CSV, encoding="utf-8-sig")
rob = pd.read_csv(ROB_CSV, encoding="utf-8-sig")
N_obs = len(df); N_firm = df["證券代碼"].nunique()


def cstar(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""


# ---------------- 表格建構 ----------------
def desc_table(doc):
    rows = [["變數", "N", "Mean", "SD", "Min", "P25", "Median", "P75", "Max"]]
    for col, lab in DESC:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        rows.append([lab, f"{s.size:,}", f"{s.mean():.3f}", f"{s.std():.3f}",
                     f"{s.min():.3f}", f"{s.quantile(.25):.3f}", f"{s.median():.3f}",
                     f"{s.quantile(.75):.3f}", f"{s.max():.3f}"])
    tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tbl.style = "Normal Table"
    for j, v in enumerate(rows[0]):
        fill(tbl.rows[0].cells[j], v, 10, True)
    for i in range(1, len(rows)):
        for j, v in enumerate(rows[i]):
            fill(tbl.rows[i].cells[j], v, 9, False, "center" if j else "left")
    three_line(tbl)


def corr_table(doc):
    cols = [c for c, _ in CORR]
    labs = [l for _, l in CORR]
    sub = df[cols].apply(pd.to_numeric, errors="coerce")
    cm = sub.corr()
    n = len(sub.dropna())
    head = ["變數"] + [f"({k+1})" for k in range(len(cols))]
    tbl = doc.add_table(rows=len(cols) + 1, cols=len(cols) + 1)
    tbl.style = "Normal Table"
    for j, v in enumerate(head):
        fill(tbl.rows[0].cells[j], v, 9, True)
    for i, lab in enumerate(labs):
        fill(tbl.rows[i + 1].cells[0], f"({i+1}) {lab}", 9, False, "left")
        for j in range(len(cols)):
            if j > i:
                fill(tbl.rows[i + 1].cells[j + 1], "", 9)
                continue
            r = cm.iloc[i, j]
            if i == j:
                fill(tbl.rows[i + 1].cells[j + 1], "1.000", 9)
            else:
                t = abs(r) * np.sqrt(max(n - 2, 1) / max(1e-9, 1 - r * r))
                fill(tbl.rows[i + 1].cells[j + 1], f"{r:.3f}", 9, bold=(t >= 1.96))
    three_line(tbl)


ROLE_ROWS = [
    ("β1", "ΔLNREV", "dLNREV"),
    ("β2", "D×ΔLNREV", "D_x_dREV"),
    ("β3", "環境變數×D×ΔLNREV", None),   # 依模型取三重交乘
    ("β4", "環境變數", None),            # 依模型取主效果
    ("", "Size×D×ΔLNREV", "Size_D_dREV"),
    ("", "AI×D×ΔLNREV", "AI_D_dREV"),
    ("", "EI×D×ΔLNREV", "EI_D_dREV"),
    ("", "ROA×D×ΔLNREV", "ROA_D_dREV"),
    ("", "Lev×D×ΔLNREV", "Lev_D_dREV"),
    ("", "Decrease×D×ΔLNREV", "Decrease_D_dREV"),
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
    s = "" if pd.isna(s) else str(s)
    return f"{b:.4f}{s}", f"({t:.2f})"


def reg_main_table(doc):
    ncol = 1 + len(L0_MODELS)
    # 每個變數兩列（係數、t）＋表頭1 ＋ N ＋ adjR2
    body_rows = len(ROLE_ROWS) * 2
    tbl = doc.add_table(rows=1 + body_rows + 2, cols=ncol)
    tbl.style = "Normal Table"
    # 表頭
    fill(tbl.rows[0].cells[0], "變數", 10, True)
    for j, (_, lab) in enumerate(L0_MODELS):
        fill(tbl.rows[0].cells[j + 1], lab, 9, True)
    ri = 1
    for greek, label, var in ROLE_ROWS:
        role = {"β1": "β1", "β2": "β2(黏性)", "β3": "β3(三重交乘)",
                "β4": "β4(水主效果)"}.get(greek)
        disp = f"{label}" + (f" ({greek})" if greek else "")
        fill(tbl.rows[ri].cells[0], disp, 9, False, "left")
        for j, (m, _) in enumerate(L0_MODELS):
            b, t = _coef_cell(m, role=role, var=var)
            fill(tbl.rows[ri].cells[j + 1], b, 9)
            fill(tbl.rows[ri + 1].cells[j + 1], t, 8)
        ri += 2
    # N
    fill(tbl.rows[ri].cells[0], "N", 9, False, "left")
    for j, (m, _) in enumerate(L0_MODELS):
        r = coef[coef["模型"] == m]
        fill(tbl.rows[ri].cells[j + 1], f"{int(r['N'].iloc[0]):,}" if len(r) else "", 9)
    ri += 1
    fill(tbl.rows[ri].cells[0], "adj. R²", 9, False, "left")
    for j, (m, _) in enumerate(L0_MODELS):
        r = coef[coef["模型"] == m]
        fill(tbl.rows[ri].cells[j + 1], f"{r['adj_R2'].iloc[0]:.4f}" if len(r) else "", 9)
    three_line(tbl)


def lag_beta3_table(doc):
    b3 = coef[coef["角色"] == "β3(三重交乘)"]
    kinds = [("主分析", "模型1\n水回收率%"), ("次分析", "模型2\n水揭露"),
             ("廢棄物主分析", "模型3\n廢棄物密集度"),
             ("廢棄物次分析", "模型4\n廢棄物揭露"), ("廢棄物穩健", "模型5\n廢棄物罰鍰")]
    lags = sorted(b3["遞延期"].unique())
    tbl = doc.add_table(rows=1 + len(lags), cols=1 + len(kinds))
    tbl.style = "Normal Table"
    fill(tbl.rows[0].cells[0], "遞延期", 9, True)
    for j, (_, lab) in enumerate(kinds):
        fill(tbl.rows[0].cells[j + 1], lab, 9, True)
    for i, lg in enumerate(lags):
        fill(tbl.rows[i + 1].cells[0], f"t-{lg}", 9, False, "left")
        for j, (kind, _) in enumerate(kinds):
            r = b3[(b3["遞延期"] == lg) & (b3["類型"] == kind)]
            if r.empty:
                fill(tbl.rows[i + 1].cells[j + 1], "", 9)
            else:
                b = r["係數"].iloc[0]; s = r["顯著性"].iloc[0]
                s = "" if pd.isna(s) else str(s)
                fill(tbl.rows[i + 1].cells[j + 1], f"{b:.4f}{s}", 9)
    three_line(tbl)


def hetero_table(doc):
    cols = ["衡量", "產業組", "遞延期", "β2", "β3", "N"]
    tbl = doc.add_table(rows=1 + len(rob), cols=len(cols))
    tbl.style = "Normal Table"
    for j, v in enumerate(cols):
        fill(tbl.rows[0].cells[j], v, 9, True)
    for i, (_, r) in enumerate(rob.iterrows()):
        b2 = "" if pd.isna(r["β2(D×ΔREV)"]) else f"{r['β2(D×ΔREV)']:.4f}{'' if r['β2_sig']=='n.s.' else r['β2_sig']}"
        b3 = "" if pd.isna(r["β3(Water×D×ΔREV)"]) else f"{r['β3(Water×D×ΔREV)']:.4f}{'' if r['β3_sig'] in ('n.s.','樣本不足') else r['β3_sig']}"
        vals = [r["水衡量"], r["產業組"], r["遞延期"], b2, b3, f"{int(r['N']):,}"]
        for j, v in enumerate(vals):
            fill(tbl.rows[i + 1].cells[j], v, 8, False, "center" if j >= 3 else "left")
    three_line(tbl)


def var_def_table(doc):
    rows = [
        ("ΔLNSGA", "LN(本期營業費用) − LN(前期營業費用)"),
        ("ΔLNREV", "LN(本期營業收入淨額) − LN(前期營業收入淨額)"),
        ("D", "收入下降虛擬：本期營收<前期營收=1，否則0"),
        ("Water_Rate", "水回收率%（TESG 企業用水量）"),
        ("Water_Disc", "水揭露虛擬：當年有水資料紀錄=1"),
        ("Waste_Intensity", "每百萬營收廢棄物量（TESG 廢棄物揭露）"),
        ("Waste_Disc", "GRI 廢棄物管理揭露度（揭露品質）"),
        ("Waste_Fine", "事業廢棄物罰鍰次數"),
        ("Size", "LN(資產總額)"),
        ("AI", "資產密集度＝資產總額/營業收入淨額"),
        ("EI", "員工密集度＝員工人數/營業收入淨額"),
        ("ROA", "資產報酬率 ROA(A)稅後息前"),
        ("Lev", "財務槓桿＝負債比率"),
        ("Decrease", "連續兩年營收下降虛擬"),
    ]
    tbl = doc.add_table(rows=len(rows) + 1, cols=2)
    tbl.style = "Normal Table"
    fill(tbl.rows[0].cells[0], "變數名稱", 10, True)
    fill(tbl.rows[0].cells[1], "定義", 10, True)
    for i, (a, b) in enumerate(rows):
        fill(tbl.rows[i + 1].cells[0], a, 9, False, "left")
        fill(tbl.rows[i + 1].cells[1], b, 9, False, "left")
    three_line(tbl)


# ---------------- 產生文件 ----------------
def main():
    doc = Document()
    n = doc.styles["Normal"]
    n.font.name = EN; n.font.size = Pt(12)
    n.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

    # 封面
    center(doc, "國立中山大學財務管理學系", 18)
    center(doc, "碩士論文", 18)
    center(doc, "Department of Finance", 16)
    center(doc, "National Sun Yat-sen University", 16)
    center(doc, "Master’s Thesis", 16)
    doc.add_paragraph()
    center(doc, "企業水管理與廢棄物管理對成本黏性之影響", 18)
    center(doc, "——以台灣上市櫃製造業之環境績效與資訊揭露為例", 14)
    center(doc, "The Effect of Corporate Water and Waste Management on Cost Stickiness", 14)
    doc.add_paragraph()
    center(doc, "研究生：（撰）", 16)
    center(doc, "指導教授：　　博士", 16)
    center(doc, "中華民國115年6月", 16)
    doc.add_page_break()

    # 摘要
    H1(doc, "摘要")
    body(doc,
         "本研究探討企業水管理與廢棄物管理之環境績效與資訊揭露，對成本黏性（cost "
         "stickiness）的影響。以台灣上市櫃製造業（排除金融保險業）"
         f"2014至2024年資料為樣本，共 {N_obs:,} 個公司—年觀測值（{N_firm:,} 家公司）。"
         "本研究採 Anderson, Banker and Janakiraman（2003）成本黏性模型，於營業費用"
         "變動對營收變動之敏感度中，加入收入下降虛擬變數與環境管理變數之三重交乘，"
         "並控制產業與年份固定效果、以公司叢集穩健標準誤估計。實證結果顯示：樣本整體"
         "存在顯著成本黏性；水管理指標之調節效果不顯著，惟廢棄物管理指標（每百萬營收"
         "廢棄物、事業廢棄物罰鍰）在高污染產業並考量時間落差後，對成本黏性呈顯著的"
         "加劇效果，支持廢棄物處理費用僵固與環境違規降本阻力之推論。")
    body(doc, "關鍵詞：成本黏性、水管理、廢棄物管理、環境資訊揭露、固定效果模型", bold=True)

    H1(doc, "Abstract")
    body(doc,
         "This study examines how corporate water and waste management—both "
         "environmental performance and information disclosure—affect cost stickiness. "
         f"Using Taiwanese listed manufacturing firms (excluding financials) over "
         f"2014-2024 ({N_obs:,} firm-year observations), we adopt the Anderson, Banker "
         "and Janakiraman (2003) framework and add triple interactions of a "
         "revenue-decrease dummy with environmental variables, controlling for industry "
         "and year fixed effects with firm-clustered robust standard errors. We find "
         "significant overall cost stickiness. Water measures show no significant "
         "moderation, whereas waste-management measures (waste per million revenue and "
         "waste-related fines) significantly intensify cost stickiness among "
         "high-pollution industries under lagged specifications.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    body(doc, "Keywords: cost stickiness, water management, waste management, "
              "environmental disclosure, fixed-effects model", bold=True)
    doc.add_page_break()

    # 緒論
    H1(doc, "緒論")
    body(doc, "在氣候變遷與淨零轉型趨勢下，水資源與廢棄物已成為企業環境管理的核心議題。"
              "此類環境投入多屬長期、專用且不可逆之資本支出與委外合約，當企業面臨營收"
              "下降時，難以即時削減，可能改變其成本調整行為。成本黏性描述成本於營收下降"
              "時向下調整幅度小於營收上升時向上調整幅度的非對稱現象，其根源在於管理者的"
              "資源調整決策與調整成本。本研究以台灣製造業為對象，檢視水與廢棄物之環境"
              "績效與資訊揭露如何影響成本黏性。")
    body(doc, "本研究貢獻有三：其一，將環境構面細分為「水資源」與「廢棄物」之實質績效與"
              "資訊揭露，較既有以整體 ESG 衡量者更為精細；其二，以成本黏性為應變數，"
              "結合時間落差與高／低污染產業異質性檢驗；其三，提供台灣製造業之實證證據，"
              "並發現廢棄物管理較水管理更具解釋力。")

    # 文獻回顧與研究假說
    H1(doc, "文獻回顧與研究假說")
    H2(doc, "文獻回顧")
    body(doc, "成本黏性文獻自 Anderson et al.（2003）以降，普遍以調整成本與管理者決策"
              "解釋成本之非對稱行為。晚近研究將環境與永續因素納入：Zeng, Peng and Chan "
              "以低碳城市政策為外生衝擊，發現綠色創新與融資限制會提高成本僵固性；"
              "Jiang and Yang（2024）指出 ESG 揭露可透過降低管理者樂觀預期而緩解供應商"
              "成本僵固性。國內方面，張凱瑜（2025）、王惠洳（2025）、陳思婷（2025）"
              "等以台灣上市櫃公司檢視成本僵固性之決定因素，辛珏辰（2025）進一步將 ESG "
              "活動納入分析，惟多以整體 ESG 為衡量，尚未細分水與廢棄物構面。")
    H2(doc, "研究假說")
    body(doc, "基於調整成本與資訊透明度兩機制，本研究提出下列假說：")
    for h in ["H1：水回收率越高，成本黏性越大（β3 預期為負）。",
              "H2：有揭露水管理資訊者，成本黏性較低（β3 預期為正）。",
              "H3：每百萬營收廢棄物越高，成本黏性越大（β3 預期為負）。",
              "H4：廢棄物管理揭露品質越高，成本黏性較低（β3 預期為正）。",
              "H5：事業廢棄物罰鍰次數越多，成本黏性越大（β3 預期為負）。"]:
        body(doc, h, indent=False, bold=True)

    # 資料描述與研究方法
    H1(doc, "資料描述與研究方法")
    H2(doc, "資料來源")
    body(doc, "本研究資料取自台灣經濟新報（TEJ），包含 IFRS 合併財務（單季，彙總為年度）、"
              "TESG 企業用水量、TESG 廢棄物揭露與罰鍰、以及 TEJ Company DB 產業別。"
              "樣本為台灣上市櫃製造業、排除金融保險業，期間 2014 至 2024 年，"
              f"經清理後共 {N_obs:,} 個公司—年觀測值、{N_firm:,} 家公司。財務單季資料以"
              "流量四季加總、存量取年末方式彙總為年度面板。")
    H2(doc, "模型設計")
    body(doc, "本研究採 ABJ 成本黏性模型，以 Env_Var 泛指環境管理變數，設定如下：")
    body(doc, "ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Env_Var×D×ΔLNREV) "
              "+ β4·Env_Var + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε", indent=False)
    body(doc, "β2 捕捉整體成本黏性（預期為負）；β3 為核心係數。Env_Var 於五個模型分別"
              "代入：模型1 水回收率%、模型2 水揭露、模型3 每百萬營收廢棄物、模型4 廢棄物"
              "揭露度、模型5 事業廢棄物罰鍰次數。並以遞延一至兩期（t-1、t-2）及依 TSE "
              "產業別分高／低污染子樣本進行檢驗；估計採 OLS 併產業與年份固定效果，"
              "並以公司層級叢集穩健標準誤。連續變數採前後 1% 縮尾。")
    H2(doc, "變數說明")
    H3(doc, "被解釋變數")
    body(doc, "以營業費用變動（ΔLNSGA）衡量成本調整行為。")
    H3(doc, "解釋變數")
    body(doc, "水管理以水回收率%（績效）與水揭露（透明度）衡量；廢棄物管理以每百萬營收"
              "廢棄物（實質投入）、GRI 廢棄物揭露度（揭露品質）與事業廢棄物罰鍰次數"
              "（違規風險）衡量。核心為環境變數與 D×ΔLNREV 之三重交乘。")
    H3(doc, "控制變數")
    body(doc, "包含企業規模（Size）、資產密集度（AI）、員工密集度（EI）、資產報酬率"
              "（ROA）、財務槓桿（Lev）與連續兩年營收下降（Decrease），並與 D×ΔLNREV "
              "交乘，另控制產業與年份固定效果。")

    # 研究結果
    H1(doc, "研究結果")
    H2(doc, "敘述統計")
    body(doc, "表4-1 呈現主要變數之敘述統計，包含平均值、標準差、最小值、四分位數與"
              "最大值。")
    caption(doc, "表4-1 敘述統計")
    desc_table(doc)
    H2(doc, "相關係數")
    body(doc, "表4-2 呈現主要連續變數之相關係數矩陣，粗體代表達 5% 顯著水準。")
    caption(doc, "表4-2 相關係數矩陣")
    corr_table(doc)
    H2(doc, "主測結果")
    b2m1 = coef[(coef["模型"] == "模型1 水回收率% L0(當期)") & (coef["角色"] == "β2(黏性)")]
    b2v = b2m1["係數"].iloc[0] if len(b2m1) else float("nan")
    body(doc, "表4-3 為五個模型之當期迴歸結果，係數下方括號為叢集穩健 t 值，*、**、*** "
              f"分別代表 10%、5%、1% 顯著水準。整體成本黏性（D×ΔLNREV，β2）於水回收率與"
              f"廢棄物密集度樣本顯著為負（如模型1 為 {b2v:.4f}），確認成本黏性存在；"
              "核心三重交乘（β3）於當期多不顯著。")
    caption(doc, "表4-3 成本黏性主迴歸結果（當期，模型1至模型5）")
    reg_main_table(doc)
    H2(doc, "延伸分析")
    body(doc, "表4-4 呈現各模型於當期（t-0）、遞延一期（t-1）與兩期（t-2）之核心係數 β3。"
              "廢棄物指標於遞延期之效果較當期明顯，支持環境投入效益需時間發酵之推論。")
    caption(doc, "表4-4 時間落差效應：各遞延期之 β3")
    lag_beta3_table(doc)
    body(doc, "表4-5 進一步依 TSE 產業別將樣本分為高耗水（高污染）與低耗水兩組，"
              "比較各衡量、產業組與遞延期之 β2 與 β3。結果顯示廢棄物密集度與罰鍰對成本"
              "黏性之加劇效果集中於高污染產業並具遞延性。")
    caption(doc, "表4-5 產業異質性 × 時間落差：β2 與 β3 對照")
    hetero_table(doc)

    # 結論
    H1(doc, "結論")
    body(doc, "本研究以台灣製造業 2014–2024 年資料，檢視水與廢棄物管理對成本黏性之影響。"
              "實證顯示樣本存在穩健成本黏性；水管理之調節效果不顯著，主因水資料樣本較小、"
              "變異有限；廢棄物管理則在高污染產業並考量時間落差後，對成本黏性呈顯著加劇"
              "效果，符合廢棄物處理委外合約僵固與環境違規降本阻力之情境。建議後續研究"
              "強化環境衡量、延長樣本期間，並以環境事件作為識別策略。")

    # 參考文獻
    H1(doc, "參考文獻")
    body(doc, "一、主要參考文獻", indent=False, bold=True)
    for r in [
        "Anderson, M. C., Banker, R. D., & Janakiraman, S. N. (2003). Are selling, "
        "general, and administrative costs “sticky”? Journal of Accounting Research, 41(1), 47–63.",
        "Zeng, J., Peng, M., & Chan, K. C. The impact of low-carbon city policy on "
        "corporate cost stickiness.",
        "Jiang, W., & Yang, W. (2024). ESG disclosure and corporate cost stickiness: "
        "Evidence from supply-chain relationships. Economics Letters, 238, 111697.",
    ]:
        body(doc, r, indent=False)
    body(doc, "二、次要參考文獻（國內相關碩博士論文，依年份新到舊）", indent=False, bold=True)
    subs = [
        (2026, "李亞峮", "ESG負面事件與媒體報導對ESG績效與盈餘管理間之關聯的調節效果", "國立臺北大學會計學系碩士論文"),
        (2026, "劉翠山", "人力資本、ESG績效(S)、ESG重大性揭露與企業市場價值關係之調節型中介模型探討", "銘傳大學會計學系碩士論文"),
        (2026, "吳東蓄", "ESG績效對代理成本之影響", "靜宜大學會計學系碩士論文"),
        (2026, "戴郁樺", "ESG投入、金融創新與綠色創新對銀行風險與績效管理之影響", "朝陽科技大學財務金融系碩士論文"),
        (2026, "楊宜蓁", "ESG及負債比率對企業信用風險之關聯性研究——以台灣上市電子業為例", "國立彰化師範大學財務金融技術學系碩士論文"),
        (2026, "楊珆佳", "臺灣金控公司ESG、金融健全指標與巴塞爾協定風險性之研究", "靜宜大學會計學系碩士論文"),
        (2026, "富得運", "ESG風險管理與企業財務績效：新興市場中企業規模之調節效果", "正修科技大學工業工程與管理碩士班碩士論文"),
        (2026, "毛祚祥", "探討台灣上市櫃觀光餐旅類股企業ESG指標、系統性風險與財務績效之關係", "國立高雄科技大學觀光管理系碩士論文"),
        (2026, "陳俊佑", "ESG績效與盈餘的價值攸關性之影響", "國防大學管理學院財務管理學系碩士論文"),
        (2026, "許嘉芸", "企業資源密集度管理對財務績效與ESG效率的影響", "國立陽明交通大學經營管理研究所碩士論文"),
        (2026, "楊博勝", "ESG資訊之揭露對於企業特有風險之影響：以台灣上市櫃公司為例", "國立陽明交通大學經營管理研究所博士論文"),
        (2026, "蘇姿云", "ESG對盈餘管理的抑制效果：家族企業的差異分析", "國立成功大學財務金融研究所碩士論文"),
        (2026, "張岑華", "ESG績效與漂綠對權益資金成本之影響：以台灣為例", "世新大學財務金融學系碩士論文"),
        (2026, "張筱婕", "漂綠行為對企業ESG之影響：媒體揭露的調節效果", "世新大學財務金融學系碩士論文"),
        (2025, "張凱瑜", "成本僵固性與企業風險關聯性之研究——以台灣上市櫃公司為例", "朝陽科技大學會計系碩士論文"),
        (2025, "王惠洳", "獨立董事連結關係與成本僵固性之關聯", "國立臺中科技大學會計資訊系碩士論文"),
        (2025, "陳思婷", "移轉訂價操弄與成本僵固性間之關聯性", "天主教輔仁大學會計學系碩士論文"),
        (2025, "辛珏辰", "ESG活動與成本僵固性——兼論COVID-19的影響", "國立成功大學會計學系碩士論文"),
        (2025, "張湘泓", "企業併購時商定留用權之勞工權益強化——由ESG觀點出發", "國立臺灣大學法律學系碩士論文"),
    ]
    for i, (yr, a, t, org) in enumerate(sorted(subs, key=lambda x: -x[0]), 1):
        body(doc, f"[{i}] {a}（{yr}）。{t}。{org}。", indent=False)

    # 附錄
    H1(doc, "附錄：變數定義")
    var_def_table(doc)

    doc.save(OUT)
    print("已生成：", OUT)


if __name__ == "__main__":
    main()
