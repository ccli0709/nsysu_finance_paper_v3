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
    ("模型5 廢棄物裁罰金額 L0(當期)", "模型5\n廢棄物裁罰金額"),
    ("模型6 用水密集度 L0(當期)", "模型6\n用水密集度"),
]
DESC = [("Y_dLNSGA", "ΔLNSGA"), ("dLNREV", "ΔLNREV"), ("D", "D"),
        ("Water_Intensity", "用水密集度"), ("Waste_Intensity", "廢棄物密集度"),
        ("Water_Rate", "水回收率%"), ("Waste_Disc", "廢棄物揭露度"),
        ("Waste_Fine", "廢棄物裁罰金額"),
        ("Size", "Size"), ("AI", "AI"), ("EI", "EI"), ("ROA", "ROA"),
        ("Lev", "Lev"), ("Decrease", "Decrease")]
CORR = [("Y_dLNSGA", "ΔLNSGA"), ("dLNREV", "ΔLNREV"), ("Water_Intensity", "用水密集度"),
        ("Waste_Intensity", "廢棄物密集度"), ("Water_Rate", "水回收率%"),
        ("Size", "Size"), ("AI", "AI"), ("ROA", "ROA"), ("Lev", "Lev")]

# ---------------- 參考文獻（統一編號）----------------
MAIN_REFS = [
    ("ABJ2003", "Anderson, M. C., Banker, R. D., & Janakiraman, S. N. (2003). Are selling, "
                "general, and administrative costs “sticky”? Journal of Accounting Research, 41(1), 47–63."),
    ("Zeng", "Zeng, J., Peng, M., & Chan, K. C. The impact of low-carbon city policy on "
             "corporate cost stickiness."),
    ("JY2024", "Jiang, W., & Yang, W. (2024). ESG disclosure and corporate cost stickiness: "
               "Evidence from supply-chain relationships. Economics Letters, 238, 111697."),
]
SUBS = [
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
    (2025, "魏郁芳", "企業內部因素與外部壓力對水資源揭露之關聯性研究：以臺灣上市櫃公司為例", "國立臺灣師範大學永續管理與環境教育研究所碩士論文"),
    (2025, "鍾文淇", "影響一般廢棄物回收率成效之研究：以直轄市為例", "國立臺北大學公共行政暨政策學系碩士論文"),
    (2024, "李依潔", "環境構面永續發展指標、企業策略與成本僵固性", "中原大學會計學系碩士論文"),
    (2024, "陳樸", "台灣上市食品公司水資源永續績效評量方法之研究", "國立臺灣大學環境工程學研究所碩士論文"),
    (2024, "顏宥盈", "成本僵固性與內外部監督效果之探討", "國立臺中科技大學會計資訊系碩士論文"),
    (2024, "謝雅筑", "飛行員CEO與成本僵固性", "國立成功大學財務金融研究所碩士論文"),
    (2023, "張芳瑜", "成本僵固性對企業績效之影響——以總體經濟變數為調節變數", "國立中興大學企業管理學系碩士論文"),
    (2023, "張宸杰", "內部控制、財務危機對成本僵固性之影響", "國立臺中科技大學會計資訊系碩士論文"),
    (2022, "劉佳熒", "供應鏈關係對供應商之成本行為的影響", "亞洲大學經營管理學系博士論文"),
    (2020, "邱圓雅", "內部債與成本僵固性之關聯性", "國立臺北商業大學會計財稅碩士班碩士論文"),
    (2018, "鄭紹君", "企業社會責任活動與成本僵固性之關聯——以企業社會責任願景調節", "國立中央大學會計研究所碩士論文"),
]
_SUB_SORTED = sorted(SUBS, key=lambda x: -x[0])
CITE_NUM = {k: i for i, (k, _) in enumerate(MAIN_REFS, 1)}
for _i, (_yr, _au, _t, _o) in enumerate(_SUB_SORTED, start=len(MAIN_REFS) + 1):
    CITE_NUM[_au] = _i


def cmark(*keys):
    nums = sorted({CITE_NUM[k] for k in keys if k in CITE_NUM})
    return "".join(f"[{n}]" for n in nums)


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
    kinds = [("用水密集度", "模型6\n用水密集度"),
             ("廢棄物主分析", "模型3\n廢棄物密集度"),
             ("主分析", "模型1\n水回收率%"), ("次分析", "模型2\n水揭露"),
             ("廢棄物次分析", "模型4\n廢棄物揭露"), ("廢棄物穩健", "模型5\n廢棄物裁罰金額")]
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
         "變動對營收變動之敏感度中，加入收入下降虛擬變數與環境變數之三重交乘，"
         "並控制產業與年份固定效果、以公司叢集穩健標準誤估計。本研究以「環境使用"
         "密集度」（用水密集度與廢棄物密集度）為核心，並以水／廢棄物之資訊揭露與違規"
         "裁罰為輔。實證結果顯示：樣本整體存在顯著成本黏性；核心之使用密集度中，"
         "用水密集度於高污染產業當期對成本黏性呈顯著加劇效果（β3 約 −0.091，5% 水準，"
         "樣本涵蓋度遠高於水回收率），廢棄物密集度亦於高污染產業遞延期呈顯著加劇；"
         "相對地，資訊揭露與違規裁罰之調節效果則不穩健。研究支持「企業對環境資源之"
         "使用密集度越高、於營收下降時成本越難削減」之推論，且效果集中於高污染產業。")
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
    H2(doc, "研究背景與動機")
    body(doc, "在氣候變遷、淨零轉型與資源稀缺的多重壓力下，水資源與廢棄物已從企業社會"
              "責任的邊陲議題，逐漸成為攸關營運存續的核心風險。台灣為海島型經濟且製造業"
              "高度密集，半導體、光電、面板、化工、鋼鐵與紡織等支柱產業普遍具有高取水、"
              "高廢水與高廢棄物之特性；2021 年台灣遭逢逾半世紀最嚴重乾旱，中部科學園區"
              "一度面臨限水，凸顯水資源供給對高耗水製造業之關鍵性。與此同時，主管機關"
              "陸續要求上市櫃公司自 2023 年起分階段編製永續報告書、並加嚴廢棄物清理與"
              "環境污染之裁處，使環境資源之取得、循環、處理與法遵成本，實質嵌入企業之"
              "營運成本結構，值得從成本行為的角度深入檢視。")
    body(doc, "企業成本如何隨營收變動而調整，是管理會計長期關注的課題。Anderson, Banker "
              "and Janakiraman（2003）提出成本僵固性（cost stickiness，或稱成本黏性）"
              "概念，指出當營收下降時，營業費用向下調整的幅度小於營收上升時向上調整的"
              "幅度，呈現顯著的非對稱性；其根源在於管理者的資源調整決策與調整成本：當"
              "面臨需求下滑時，若管理者預期衰退僅屬短暫，或裁撤與重建專用資源之成本高昂，"
              "便傾向保留閒置產能與人力，形成成本黏性。" + cmark("ABJ2003"))
    body(doc, "環境資源的投入具備高投入、長週期、專用性與不可逆之特徵。企業為提升水回收、"
              "處理廢水與廢棄物，往往需建置廢水處理廠、水循環系統與專用處理設施，並簽訂"
              "長期委外清運與處理合約；這類與產能高度綁定的環境支出，於營收下降時難以"
              "即時削減，理論上將加劇成本黏性。另一方面，環境資訊之揭露被視為降低資訊"
              "不對稱、矯正管理者過度樂觀預期之機制，可能反向緩解成本黏性。因此，環境"
              "因素對成本黏性的影響，同時存在「實質投入加劇」與「資訊透明緩解」兩股方向"
              "相反的力量，值得以台灣製造業資料加以釐清。")
    body(doc, "既有將環境因素連結成本行為的研究，多以整體 ESG 評分或揭露為衡量，較少"
              "深入「水」與「廢棄物」等個別、可量化之實質環境構面，且常以「揭露有無」或"
              "「罰鍰次數」等粗略代理，易受衡量誤差與單一極端值干擾。本研究主張，真正"
              "貼近調整成本本質的，是企業對環境資源之「使用密集度」——單位營收所耗用之"
              "水量與產生之廢棄物量，直接反映其營運與環境資源之綁定程度及專用處理投入。"
              "相對於水回收率等揭露樣本稀少、變異有限之指標，用水密集度（總用水量／營收）"
              "之樣本涵蓋度更高，更能提供穩健且具檢定力之證據。"
              + cmark("JY2024", "辛珏辰", "鄭紹君"))
    body(doc, "從調整成本理論的角度，成本黏性反映管理者於需求變動時「保留或裁撤資源」之"
              "權衡。當向下調整資源（如解約、資遣、處分設施）之成本高於暫時保留閒置資源"
              "之成本，或管理者對未來需求抱持樂觀預期時，便傾向於營收下降時不立即縮減"
              "成本。環境資源之投入恰具使黏性加劇之多重特徵：其一，廢水與廢棄物處理設施"
              "屬專用性資產，缺乏次級市場、處分價值低；其二，環境委外處理多採長期合約，"
              "短期難以中止；其三，環境法遵具強制性，企業不能為節省成本而停止污染防治。"
              "因此，企業對環境資源之使用密集度越高，理應於營收下降時面臨越強之向下調整"
              "阻力，此即本研究之核心邏輯。")
    body(doc, "進一步言，台灣製造業在全球供應鏈中扮演關鍵角色，其生產活動高度倚賴穩定"
              "且充足之水電與原物料供給，同時受環境法規日趨嚴格之規範。近年制度快速"
              "演進：缺水風險促使高耗水產業投入再生水與製程節水；廢棄物清理與資源循環"
              "規範要求企業建立妥善之清運、處理與再利用機制；污染裁處與環境違規之公告，"
              "亦使企業之環境法遵成本與商譽風險同步上升。這些制度變遷意味企業在環境面之"
              "投入與承諾具長期性與剛性，一旦建置便難以隨短期營收波動而快速調整，正是"
              "檢視環境因素如何形塑成本黏性之合適場域。")
    body(doc, "本研究之所以將「水」與「廢棄物」並列檢視，係因二者同為製造業最主要之環境"
              "資源投入與產出構面，且皆具備專用設施、長期合約與強制法遵之特性，在成本"
              "結構上具高度可比性；同時，兩者在 TEJ 資料庫中皆有可量化之使用量、密集度、"
              "揭露與裁罰資料，便於建立一致之衡量與比較架構。並列水與廢棄物有助於檢視"
              "環境使用密集度效果之普遍性與穩健性，並辨識何種構面、何種產業與何種時點下，"
              "環境投入對成本行為之影響最為顯著。")
    body(doc, "成本黏性不僅是學術議題，更具實務意涵：成本向下調整之僵固性會影響企業於"
              "景氣下行時之獲利韌性、盈餘品質與現金流量，並牽動分析師預測與投資人評價。"
              "理解哪些因素會強化或緩解成本黏性，有助於評估企業於營運波動下之調整彈性；"
              "將環境資源投入納入此一框架，則可回應永續轉型下「環境承諾是否影響企業財務"
              "韌性」之關切，兼具理論與政策意義。")
    body(doc, "就台灣製造業而言，半導體與光電之超純水需求、化工與鋼鐵之大量製程用水與"
              "廢水、紡織染整之高污染負荷，均使水與廢棄物之處理成為顯著且持續之營運支出；"
              "加以再生水建置、廢棄物再利用與污染防治設備多屬長期資本投入，形成難以逆轉"
              "之成本承諾。此一產業特性使台灣成為檢視環境使用密集度與成本黏性關係之理想"
              "場域，也使本研究之發現對高污染產業之成本管理與韌性評估具參考價值。")
    body(doc, "2021 年之乾旱使部分高耗水廠區面臨限水與購水成本上升，凸顯用水密集度高之"
              "企業於資源供給衝擊下之脆弱性；此類衝擊雖屬短期，卻反映企業對水資源之"
              "結構性依賴。將此依賴以「用水密集度」量化，並連結至企業於營收下降時之成本"
              "調整行為，正是本研究有別於既有以整體 ESG 或揭露為衡量之研究的切入點。")
    body(doc, "為回答上述問題，本研究以台灣經濟新報（TEJ）之財務、TESG 企業用水量、"
              "廢棄物揭露與列管事業污染源裁處等資料，建構 2014 至 2024 年台灣上市櫃"
              "製造業之公司—年面板，採 ABJ 成本僵固性模型，將用水密集度、廢棄物密集度"
              "及其揭露與裁罰變數，與收入下降虛擬變數交乘，並控制產業與年份固定效果、"
              "採公司叢集穩健標準誤；另以時間落差（遞延一至兩期）與高／低污染產業異質性，"
              "檢驗環境使用密集度對成本黏性之影響及其適用情境。此一設計兼顧衡量效度、"
              "統計檢定力與因果詮釋之層次，為本研究之方法特色。")
    H2(doc, "研究目的與研究問題")
    body(doc, "本研究之研究目的有四：第一，建立以「環境使用密集度」為核心之成本黏性分析"
              "架構，檢視用水密集度與廢棄物密集度是否顯著影響台灣製造業之成本黏性；"
              "第二，以資訊揭露（水揭露、廢棄物揭露品質）與違規裁罰（廢棄物裁罰金額）為"
              "輔助構面，比較「實質使用」與「資訊透明度」兩類機制之相對解釋力；第三，"
              "考量環境投入之效益與僵固性需時間發酵，檢驗當期與遞延（t-1、t-2）之差異；"
              "第四，考量台灣製造業次產業之環境依賴度差異甚大，比較高污染與低污染產業之"
              "異質性，釐清效果集中之情境。本研究範圍限縮於台灣上市櫃製造業並排除金融"
              "保險業，樣本期間為 2014 至 2024 年。")
    body(doc, "本研究之預期貢獻可分三方面。理論上，將環境因素對成本行為之影響，由抽象之"
              "整體 ESG 或 CSR 評分，推進至可量化、可比較之「實質使用密集度」，使調整"
              "成本理論在環境情境下獲得更貼近本質之檢驗。衡量上，本研究以連續之「用水"
              "密集度」與「廢棄物裁罰金額」取代樣本稀少之水回收率與易受極端值主導之罰鍰"
              "次數，兼顧樣本涵蓋度與衡量效度。方法上，結合時間落差與高／低污染產業異質"
              "性，揭示環境成本效應具「產業依存與遞延顯現」之特徵，為僅以全樣本、當期"
              "估計之研究所難以觀察。本文後續安排：第二章文獻回顧與假說、第三章資料與"
              "研究方法、第四章實證結果、第五章結論與建議。")

    # 文獻回顧與研究假說
    H1(doc, "文獻回顧與研究假說")
    H2(doc, "文獻回顧")
    body(doc, "（一）成本僵固性之理論根源與決定因素。成本僵固性研究以 Anderson, Banker "
              "and Janakiraman（2003）為濫觴，其以美國公司之銷售、一般及管理費用（SG&A）"
              "驗證營收下降時費用調整之非對稱性，並提出「調整成本」與「管理者樂觀預期」"
              "兩大解釋機制。後續文獻延伸至多元決定因素，國內以台灣上市櫃公司為樣本者"
              "尤眾：張凱瑜（2025）發現成本僵固性與企業風險相關；王惠洳（2025）指出"
              "獨立董事之連結關係會影響成本僵固性；陳思婷（2025）以移轉訂價操弄反映代理"
              "問題，發現操弄程度越高、成本僵固性越大；張宸杰（2023）檢視內部控制與財務"
              "危機之影響；顏宥盈（2024）探討內外部監督效果；謝雅筑（2024）以「飛行員"
              "CEO」之風險偏好特質、邱圓雅（2020）以內部債，分別檢視管理者誘因；張芳瑜"
              "（2023）則反向檢視成本僵固性對企業績效之影響，並以總體經濟變數為調節。"
              "上述研究共同確立成本僵固性在台灣市場穩健存在，且受治理、誘因、財務狀況與"
              "總體環境調節，構成本研究之方法論基礎。"
              + cmark("ABJ2003", "張凱瑜", "王惠洳", "陳思婷", "張宸杰", "顏宥盈",
                      "謝雅筑", "邱圓雅", "張芳瑜"))
    body(doc, "（二）環境、ESG 與成本行為。將環境與永續因素連結成本行為之研究漸增。"
              "Zeng, Peng and Chan 以中國低碳城市政策為外生衝擊，發現政策促使企業增加"
              "綠色創新投資、並受融資限制影響，進而提高成本僵固性；Jiang and Yang（2024）"
              "指出客戶端之 ESG 揭露可透過供應鏈關係降低供應商管理者之過度樂觀，緩解其"
              "成本僵固性。國內方面，鄭紹君（2018）檢視企業社會責任活動與成本僵固性之"
              "關聯並以 CSR 願景為調節；最具參考價值者為李依潔（2024），其以 2016 至 "
              "2022 年台灣上市櫃公司為樣本，發現「企業投入 ESG 環境活動績效越高、成本"
              "僵固性越高」，且採探勘者策略之公司因技術彈性而減緩此正向關係——與本研究"
              "「環境實質投入加劇成本黏性」之推論一致；辛珏辰（2025）進一步將 ESG 活動與"
              "COVID-19 衝擊納入分析。惟上述研究多以整體 ESG 或 CSR 為衡量，尚未細分水與"
              "廢棄物等個別環境構面。"
              + cmark("Zeng", "JY2024", "鄭紹君", "李依潔", "辛珏辰"))
    body(doc, "（三）水資源績效與揭露。陳樸（2024）針對台灣上市食品公司建構水資源永續"
              "績效之評量方法，凸顯不同產業之用水特性與績效衡量之重要性；魏郁芳（2025）"
              "檢視企業內部因素與外部壓力對水資源揭露之影響，指出揭露行為受規範壓力與"
              "公司特徵驅動。這些研究確立水資源之績效與揭露可被量化、且具明顯產業異質性，"
              "但少有將其進一步連結至企業成本調整行為者，留下可供延伸之空間。值得注意"
              "的是，水回收率等揭露型指標之樣本涵蓋度偏低、變異有限，若以其為主分析易"
              "受樣本選擇與檢定力不足之限制；相對地，以「總用水量／營收」建構之用水"
              "密集度可涵蓋更廣之樣本、且直接反映營運對水資源之依賴強度，較適合作為"
              "核心衡量。此一衡量選擇之考量，構成本研究以使用密集度為核心之重要理由。"
              + cmark("陳樸", "魏郁芳"))
    body(doc, "（四）廢棄物管理與環境裁罰。鍾文淇（2025）以直轄市為對象研究一般廢棄物"
              "回收率之成效，反映廢棄物處理之政策與制度脈絡。就企業而言，廢棄物之清運、"
              "處理與再利用高度仰賴專用設施與長期委外合約，且環境違規將面臨主管機關裁處，"
              "形成法遵與改善之強制性支出。相較於僅計算「罰鍰次數」，TEJ 之列管事業污染源"
              "裁處資料提供逐案「裁處金額」，更能衡量違規之經濟強度，並可避免計數型變數"
              "受單一極端值主導之偏誤。廢棄物構面同時具「實質產出密集度」與「違規風險」"
              "兩面向，本研究將兩者分別檢視。尤以有害廢棄物之清運與處理受法規嚴格規範、"
              "單位處理成本高昂且不可省略，其密集度越高者，於營收下降時越難壓縮相關支出，"
              "理論上對成本黏性之推升作用更強。" + cmark("鍾文淇"))
    body(doc, "（五）供應鏈機制與研究缺口。劉佳熒（2022）檢視客戶—供應商關係對供應商"
              "成本行為之影響，與 Jiang and Yang（2024）相互呼應，顯示外部關係亦形塑"
              "企業之成本調整。此外，Anderson et al.（2003）之對數—對數設定已成為成本"
              "僵固性研究之標準，便於加入三重交乘以檢視特定因素之調節，國內文獻多沿用之，"
              "本研究亦以之為基礎。綜合而言，既有文獻雖已分別累積「成本僵固性決定因素」"
              "與「環境／ESG 財務後果」兩支脈絡，但少有研究以「成本黏性」為應變數，並將"
              "環境構面細分為水與廢棄物之「實質使用密集度」與「資訊揭露」，且同時結合"
              "時間落差與高／低污染產業異質性加以檢驗，此即本研究欲填補之缺口。"
              + cmark("劉佳熒", "JY2024", "ABJ2003"))
    body(doc, "（六）理論整合與本研究之定位。統整上述文獻，環境因素影響成本黏性主要循"
              "兩條路徑：其一為「調整成本／實質投入」路徑，企業投入水循環與廢棄物處理等"
              "專用資產與長期合約後，於營收下降時難以即時削減，遂加劇成本黏性（Zeng et "
              "al.、李依潔 2024 屬之）；其二為「資訊透明度／監督」路徑，高品質環境揭露"
              "降低資訊不對稱、矯正管理者過度樂觀，促使其果斷調整資源而緩解黏性（Jiang "
              "and Yang 2024、魏郁芳 2025 屬之）。本研究主張，就水與廢棄物而言，實質使用"
              "密集度所代表之調整成本路徑較揭露路徑更直接且可量化，故以使用密集度為核心"
              "構面、揭露與裁罰為輔助對照；並因高污染產業之環境專用投入與法遵負擔遠高於"
              "組裝與服務類產業，效果應集中於高污染產業，且隨設施折舊與合約續期而遞延"
              "顯現，故另檢視高／低污染子樣本與遞延一至兩期之效果。"
              + cmark("Zeng", "李依潔", "JY2024", "魏郁芳"))
    body(doc, "（七）研究缺口與假說對應。綜合前述，本研究之文獻定位可歸納為三點：第一，"
              "成本僵固性之存在與決定因素已獲國內外文獻充分支持，惟環境構面之探討仍以"
              "整體 ESG／CSR 為主；第二，水與廢棄物之績效與揭露雖已可量化，卻鮮少連結至"
              "成本行為；第三，既有以「揭露有無」或「罰鍰次數」之粗略衡量易生偏誤。據此，"
              "本研究以使用密集度為核心（對應核心假說 H6 及廢棄物實質投入 H3）、以資訊"
              "揭露與違規裁罰為輔（對應 H1、H2、H4、H5），並以高／低污染產業與時間落差"
              "檢驗其適用情境，冀補充上述缺口並強化因果詮釋之層次。")
    body(doc, "在衡量與模型方面，Anderson et al.（2003）之對數—對數設定已成為成本僵固性"
              "研究之標準：以營業費用變動率對營收變動率迴歸，並加入營收下降虛擬變數與"
              "營收變動率之交乘，藉交乘項係數捕捉調整之非對稱性；此設定之優點在於係數可"
              "直接解讀為彈性，且便於加入調節變數之三重交乘。國內文獻（如張凱瑜 2025、"
              "陳思婷 2025、李依潔 2024）多沿用此架構，顯示其在台灣資料之適用性；本研究"
              "亦以之為基礎，將環境使用密集度、揭露與裁罰變數納入三重交乘。此外，隨金管會"
              "分階段推動永續報告書之編製與第三方確信，企業於用水、廢水與廢棄物之量化"
              "揭露日趨完整，TESG 亦據以建置用水量、廢棄物量與 GRI 揭露度等結構化欄位，"
              "使本研究得以較過去更完整地衡量企業之環境使用密集度與揭露品質，並將列管"
              "事業污染源之逐案裁處金額納入分析，兼顧衡量之細緻度與可靠度。"
              + cmark("ABJ2003", "張凱瑜", "陳思婷", "李依潔"))
    H2(doc, "研究假說")
    body(doc, "基於調整成本與資訊透明度兩機制，本研究提出下列假說；其中 H6 為核心假說："
             , indent=False)
    for h in ["H1：水回收率越高，成本黏性越大（β3 預期為負）。",
              "H2：有揭露水管理資訊者，成本黏性較低（β3 預期為正）。",
              "H3：每百萬營收廢棄物越高，成本黏性越大（β3 預期為負）。",
              "H4：廢棄物管理揭露品質越高，成本黏性較低（β3 預期為正）。",
              "H5：廢棄物違規裁罰金額越高，成本黏性越大（β3 預期為負）。",
              "H6（核心）：企業用水密集度（與廢棄物密集度）越高，成本黏性越大，"
              "且效果集中於高污染產業（β3 預期為負）。"]:
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
    body(doc, "β2 捕捉整體成本黏性（預期為負）；β3 為核心係數。Env_Var 於六個模型分別"
              "代入：核心為模型6 用水密集度（總用水量/營收）與模型3 每百萬營收廢棄物；"
              "輔助為模型1 水回收率%、模型2 水揭露、模型4 廢棄物揭露度、模型5 廢棄物"
              "裁罰金額佔資產（取自列管事業污染源裁處明細之裁處金額，取代罰鍰次數以"
              "避免極端值主導）。並以遞延一至兩期（t-1、t-2）及依 TSE 產業別分高／低污染"
              "子樣本進行檢驗；估計採 OLS 併產業與年份固定效果，並以公司層級叢集穩健"
              "標準誤。連續變數採前後 1% 縮尾。")
    H2(doc, "變數說明")
    H3(doc, "被解釋變數")
    body(doc, "以營業費用變動（ΔLNSGA）衡量成本調整行為。")
    H3(doc, "解釋變數")
    body(doc, "核心之使用密集度以用水密集度（總用水量／營收）與廢棄物密集度（每百萬營收"
              "廢棄物）衡量；輔助構面以水回收率%、水揭露（透明度）、GRI 廢棄物揭露度"
              "（揭露品質）與廢棄物裁罰金額佔資產（違規風險）衡量。核心為環境變數與 "
              "D×ΔLNREV 之三重交乘。")
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
    body(doc, "表4-3 為六個模型之當期迴歸結果（核心：模型6 用水密集度、模型3 廢棄物"
              "密集度；輔助：模型1、2、4、5），係數下方括號為叢集穩健 t 值，*、**、*** "
              f"分別代表 10%、5%、1% 顯著水準。整體成本黏性（D×ΔLNREV，β2）於使用密集度"
              f"相關樣本顯著為負（如模型1 為 {b2v:.4f}），確認成本黏性存在；核心三重交乘"
              "（β3）於全樣本當期多不顯著，其顯著性主要顯現於高污染產業子樣本（見延伸分析）。")
    caption(doc, "表4-3 成本黏性主迴歸結果（當期，模型1至模型6）")
    reg_main_table(doc)
    H2(doc, "延伸分析")
    body(doc, "表4-4 呈現各模型於當期（t-0）、遞延一期（t-1）與兩期（t-2）之核心係數 β3。"
              "使用密集度（用水密集度、廢棄物密集度）之效果較揭露與裁罰明顯，且部分於"
              "遞延期方顯現，支持環境投入效益與僵固性需時間發酵之推論。")
    caption(doc, "表4-4 時間落差效應：各遞延期之 β3")
    lag_beta3_table(doc)
    body(doc, "表4-5 進一步依 TSE 產業別將樣本分為高耗水（高污染）與低耗水兩組，"
              "比較各衡量、產業組與遞延期之 β2 與 β3。結果顯示用水密集度於高污染產業"
              "當期（β3 約 −0.091，5% 水準）、廢棄物密集度於高污染產業遞延期，對成本"
              "黏性之加劇效果最為明顯，顯示效果集中於高污染產業且具產業依存與遞延特徵。")
    caption(doc, "表4-5 產業異質性 × 時間落差：β2 與 β3 對照")
    hetero_table(doc)

    # 結論
    H1(doc, "結論")
    body(doc, "本研究以台灣製造業 2014–2024 年資料，以「環境使用密集度」為核心，檢視水與"
              "廢棄物之實質投入與資訊揭露對成本黏性之影響。實證顯示樣本存在穩健成本黏性；"
              "核心之使用密集度中，用水密集度於高污染產業當期對成本黏性呈顯著加劇效果"
              "（β3 約 −0.091，5% 水準，且樣本涵蓋度遠高於水回收率），廢棄物密集度亦於"
              "高污染產業遞延期呈顯著加劇；相對地，水回收率、水揭露、廢棄物揭露與廢棄物"
              "裁罰金額之調節效果則不穩健。此結果支持：企業對環境資源之使用密集度越高，"
              "營運與專用處理設施及長期委外合約之綁定越深，於營收下降時成本越難即時削減，"
              "從而加劇成本黏性，且此效應集中於高污染產業並具遞延性。")
    body(doc, "本研究之貢獻在於：將環境對成本行為之影響，由抽象之整體 ESG 揭露推進至可"
              "量化之「實質使用密集度」；並指出以罰鍰次數衡量環境違規之偏誤（易受單一"
              "極端值主導），改採裁處金額佔資產後結論更為穩健；以用水密集度取代樣本稀少"
              "之水回收率，顯著提升檢定力。政策上，建議主管機關重視高污染產業使用密集度"
              "資訊之標準化與可比較性；實務上，高使用密集度企業宜於資本配置與產能規劃"
              "時將此僵固性納入考量，並透過循環化與資源效率投資降低對環境資源之剛性依賴。"
              "後續研究可延伸至碳排密集度等其他環境使用構面，或以環境稽查與缺水等外生"
              "事件強化因果識別。")

    # 參考文獻（統一編號，內文以 [n] 對應）
    H1(doc, "參考文獻")
    body(doc, "（採統一編號，內文以 [n] 對應；[1]–[3] 為核心理論與模型來源，"
              "[4] 起為國內相關碩博士論文，依年份新到舊。）", indent=False)
    for i, (_k, ref) in enumerate(MAIN_REFS, 1):
        body(doc, f"[{i}] {ref}", indent=False)
    for i, (yr, a, t, org) in enumerate(_SUB_SORTED, start=len(MAIN_REFS) + 1):
        body(doc, f"[{i}] {a}（{yr}）。{t}。{org}。", indent=False)

    # 附錄
    H1(doc, "附錄：變數定義")
    var_def_table(doc)

    doc.save(OUT)
    print("已生成：", OUT)


if __name__ == "__main__":
    main()
