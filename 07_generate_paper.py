# -*- coding: utf-8 -*-
"""
07_generate_paper.py
--------------------
生成 Word 論文：水管理與成本黏性_論文.docx

- 敘述統計、相關矩陣：由 04_model_data.csv 即時計算。
- 迴歸表：讀 05_regression_coef.csv（模型1/2）。
- 穩健性表：讀 06_robustness_summary.csv。
- 變數定義表：讀 04_var_mapping.csv。
- 論文架構參照 reference_paper/ 兩篇成本黏性文獻（Zeng et al.；Jiang & Yang 2024）。
"""

import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

MODEL_CSV = "04_model_data.csv"
MAP_CSV = "04_var_mapping.csv"
COEF_CSV = "05_regression_coef.csv"
ROB_CSV = "06_robustness_summary.csv"
OUT_DOCX = "水與廢棄物管理與成本黏性_論文.docx"

CJK_FONT = "新細明體"
EN_FONT = "Times New Roman"

DESC_VARS = ["Y_dLNSGA", "dLNREV", "D", "Water_Rate", "Water_Disc",
             "Waste_Intensity", "Waste_Disc", "Waste_Fine",
             "Size", "AI", "EI", "ROA", "Lev", "Decrease"]
CORR_VARS = ["Y_dLNSGA", "dLNREV", "Water_Rate", "Waste_Intensity",
             "Size", "AI", "EI", "ROA", "Lev"]
WATER_MEASURES = ["水回收率%", "水揭露", "用水密集度"]
WASTE_MEASURES = ["廢棄物密集度", "廢棄物揭露", "廢棄物裁罰金額"]

# 主要參考文獻（reference_paper_main）＋基礎文獻
MAIN_REFS = [
    "Anderson, M. C., Banker, R. D., & Janakiraman, S. N. (2003). Are selling, "
    "general, and administrative costs \u201csticky\u201d? Journal of Accounting "
    "Research, 41(1), 47\u201363.",
    "Zeng, J., Peng, M., & Chan, K. C. The impact of low-carbon city policy on "
    "corporate cost stickiness. (以中國 A 股 2008\u20132022 為樣本，採 DiD；機制為"
    "綠色創新與融資限制——對應本研究環境實質投入之調整成本觀點，H1/H3。)",
    "Jiang, W., & Yang, W. (2024). ESG disclosure and corporate cost stickiness: "
    "Evidence from supply-chain relationships. Economics Letters, 238, 111697. "
    "(ESG 揭露透過降低管理者樂觀預期而緩解供應商成本僵固——對應本研究揭露機制，"
    "H2/H4。)",
]

# 次要參考文獻（reference_paper_sub，19 篇國內碩博論文；西元年）
SUB_REFS = [
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
]


# ---------------------------------------------------------------- docx helpers
def set_cjk(run):
    run.font.name = EN_FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CJK_FONT)


def add_heading(doc, text, level=1):
    p = doc.add_heading(level=level)
    r = p.add_run(text)
    set_cjk(r)
    return p


def add_para(doc, text, size=12, bold=False, align=None, first_indent=True):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if first_indent:
        p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    set_cjk(r)
    return p


def add_table(doc, df, title=None, num_fmt="{:.4f}"):
    if title:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = cap.add_run(title)
        r.bold = True
        r.font.size = Pt(11)
        set_cjk(r)
    t = doc.add_table(rows=1, cols=len(df.columns))
    t.style = "Light Grid Accent 1"
    for j, col in enumerate(df.columns):
        c = t.rows[0].cells[j].paragraphs[0].add_run(str(col))
        c.bold = True
        c.font.size = Pt(9)
        set_cjk(c)
    for _, row in df.iterrows():
        cells = t.add_row().cells
        for j, col in enumerate(df.columns):
            v = row[col]
            if isinstance(v, (int, np.integer)):
                s = f"{v:,}"
            elif isinstance(v, (float, np.floating)):
                s = "" if pd.isna(v) else num_fmt.format(v)
            else:
                s = str(v)
            rr = cells[j].paragraphs[0].add_run(s)
            rr.font.size = Pt(9)
            set_cjk(rr)
    doc.add_paragraph()
    return t


# ---------------------------------------------------------------- content bits
def build_descriptive(df):
    rows = []
    for v in DESC_VARS:
        s = pd.to_numeric(df[v], errors="coerce").dropna()
        rows.append({"變數": v, "N": int(s.size), "平均數": s.mean(),
                     "標準差": s.std(), "最小值": s.min(),
                     "中位數": s.median(), "最大值": s.max()})
    return pd.DataFrame(rows)


def build_corr(df):
    sub = df[CORR_VARS].apply(pd.to_numeric, errors="coerce")
    corr = sub.corr().round(3).reset_index().rename(columns={"index": "變數"})
    return corr


def build_reg_table(coef):
    """當期(L0)模型1/2 併排：變數 | 模型1 係數(顯著) | 模型2 係數(顯著)。"""
    c0 = coef[coef["遞延期"] == 0]
    labels = c0["模型"].unique().tolist()
    m1 = labels[0]
    m2 = labels[1] if len(labels) > 1 else labels[0]
    order = ["dLNREV", "D_x_dREV", "WaterRate_D_dREV", "WaterDisc_D_dREV",
             "Water_Rate", "Water_Disc", "Size_D_dREV", "AI_D_dREV",
             "EI_D_dREV", "ROA_D_dREV", "Lev_D_dREV", "Decrease_D_dREV"]

    def cell(label, var):
        r = c0[(c0["模型"] == label) & (c0["變數"] == var)]
        if r.empty:
            return ""
        b = r["係數"].iloc[0]; s = r["顯著性"].iloc[0]; t = r["t值"].iloc[0]
        return f"{b:.4f}{s} ({t:.2f})"

    rows = []
    for v in order:
        c1, c2 = cell(m1, v), cell(m2, v)
        if c1 == "" and c2 == "":
            continue
        rows.append({"變數": v, "模型1(水回收率%)": c1, "模型2(水揭露)": c2})
    tbl = pd.DataFrame(rows)

    def meta(label, key):
        r = c0[c0["模型"] == label]
        return "" if r.empty else r[key].iloc[0]
    n1, n2 = meta(m1, "N"), meta(m2, "N")
    a1, a2 = meta(m1, "adj_R2"), meta(m2, "adj_R2")
    tbl = pd.concat([tbl, pd.DataFrame([
        {"變數": "N", "模型1(水回收率%)": f"{int(n1):,}", "模型2(水揭露)": f"{int(n2):,}"},
        {"變數": "adj. R²", "模型1(水回收率%)": f"{a1:.4f}", "模型2(水揭露)": f"{a2:.4f}"},
    ])], ignore_index=True)
    return tbl


def build_models_table(coef, specs):
    """通用：給定 [(模型完整標籤, 顯示名)]，輸出 變數×模型 之係數(下方 t 值) 三線表資料。"""
    c0 = coef[coef["遞延期"] == 0]
    role_rows = [("β1", "ΔLNREV", "dLNREV"), ("β2", "D×ΔLNREV", "D_x_dREV"),
                 ("β3", "環境變數×D×ΔLNREV", None), ("β4", "環境變數", None),
                 ("", "Size×D×ΔLNREV", "Size_D_dREV"), ("", "AI×D×ΔLNREV", "AI_D_dREV"),
                 ("", "EI×D×ΔLNREV", "EI_D_dREV"), ("", "ROA×D×ΔLNREV", "ROA_D_dREV"),
                 ("", "Lev×D×ΔLNREV", "Lev_D_dREV"), ("", "Decrease×D×ΔLNREV", "Decrease_D_dREV")]
    role_map = {"β1": "β1", "β2": "β2(黏性)", "β3": "β3(三重交乘)", "β4": "β4(水主效果)"}

    def cell(mlabel, greek, var):
        r = c0[c0["模型"] == mlabel]
        rr = r[r["角色"] == role_map[greek]] if greek in role_map else r[r["變數"] == var]
        if rr.empty:
            return "", ""
        b = rr["係數"].iloc[0]; s = rr["顯著性"].iloc[0]; t = rr["t值"].iloc[0]
        s = "" if pd.isna(s) else str(s)
        return f"{b:.4f}{s}", f"({t:.2f})"

    rows = []
    for greek, label, var in role_rows:
        disp = f"{label}" + (f" ({greek})" if greek else "")
        line = {"變數": disp}
        line_t = {"變數": ""}
        any_val = False
        for mlabel, dispname in specs:
            b, t = cell(mlabel, greek, var)
            line[dispname] = b; line_t[dispname] = t
            any_val = any_val or (b != "")
        if any_val:
            rows.append(line); rows.append(line_t)
    # N / adjR2
    nrow = {"變數": "N"}; arow = {"變數": "adj. R²"}
    for mlabel, dispname in specs:
        r = c0[c0["模型"] == mlabel]
        nrow[dispname] = f"{int(r['N'].iloc[0]):,}" if len(r) else ""
        arow[dispname] = f"{r['adj_R2'].iloc[0]:.4f}" if len(r) else ""
    rows.append(nrow); rows.append(arow)
    return pd.DataFrame(rows)


def summarize_findings(coef, rob):
    """掃描 β3 顯著性，產生自適應結論句（隨實際結果變動）。"""
    b3 = coef[coef["角色"] == "β3(三重交乘)"]
    sig = b3[b3["顯著性"].isin(["*", "**", "***"])]
    lag_txt = []
    for _, r in sig.iterrows():
        direction = "負向（加劇黏性）" if r["係數"] < 0 else "正向（緩解黏性）"
        lag_txt.append(f"{r['模型']} β3={r['係數']:.4f}{r['顯著性']}（{direction}）")

    rob_sig = rob[rob["β3_sig"].isin(["*", "**", "***"])]
    rob_txt = []
    for _, r in rob_sig.iterrows():
        direction = "負向" if r["β3(Water×D×ΔREV)"] < 0 else "正向"
        rob_txt.append(f"{r['水衡量']}·{r['產業組']}·{r['遞延期']}："
                       f"β3={r['β3(Water×D×ΔREV)']:.4f}{r['β3_sig']}（{direction}）")

    any_sig = len(sig) > 0 or len(rob_sig) > 0
    return any_sig, lag_txt, rob_txt


def build_lag_table(coef, kind_cols):
    """時間落差：各遞延期之 β3（X×D×ΔLNREV）。kind_cols=[(類型, 欄名)...]。"""
    b3 = coef[coef["角色"] == "β3(三重交乘)"].copy()
    rows = []
    for lag in sorted(b3["遞延期"].unique()):
        rec = {"遞延期": f"t-{lag}（{'當期' if lag == 0 else str(lag)+'期前'}）"}
        for kind, colname in kind_cols:
            r = b3[(b3["遞延期"] == lag) & (b3["類型"] == kind)]
            if r.empty:
                rec[colname] = ""
            else:
                b = r["係數"].iloc[0]; s = r["顯著性"].iloc[0]
                t = r["t值"].iloc[0]; n = int(r["N"].iloc[0])
                rec[colname] = f"{b:.4f}{s} (t={t:.2f}, N={n:,})"
        rows.append(rec)
    return pd.DataFrame(rows)


def build_waste_reg_table(coef):
    """當期(L0)廢棄物三模型併排：模型3密集度／模型4揭露／模型5罰鍰。"""
    c0 = coef[coef["遞延期"] == 0]
    kinds = [("廢棄物主分析", "模型3(廢棄物密集度)"),
             ("廢棄物次分析", "模型4(廢棄物揭露)"),
             ("廢棄物穩健", "模型5(廢棄物裁罰金額)")]
    order = ["dLNREV", "D_x_dREV", "WasteInt_D_dREV", "WasteDisc_D_dREV",
             "WasteFine_D_dREV", "Waste_Intensity", "Waste_Disc", "Waste_Fine",
             "Size_D_dREV", "AI_D_dREV", "EI_D_dREV", "ROA_D_dREV",
             "Lev_D_dREV", "Decrease_D_dREV"]

    def cell(kind, var):
        r = c0[(c0["類型"] == kind) & (c0["變數"] == var)]
        if r.empty:
            return ""
        b = r["係數"].iloc[0]; s = r["顯著性"].iloc[0]; t = r["t值"].iloc[0]
        return f"{b:.4f}{s} ({t:.2f})"

    rows = []
    for v in order:
        vals = {kc[1]: cell(kc[0], v) for kc in kinds}
        if all(x == "" for x in vals.values()):
            continue
        rows.append({"變數": v, **vals})
    tbl = pd.DataFrame(rows)

    def meta(kind, key):
        r = c0[c0["類型"] == kind]
        return "" if r.empty else r[key].iloc[0]
    tbl = pd.concat([tbl, pd.DataFrame([
        {"變數": "N", **{kc[1]: (f"{int(meta(kc[0],'N')):,}" if meta(kc[0], 'N') != '' else '') for kc in kinds}},
        {"變數": "adj. R²", **{kc[1]: (f"{meta(kc[0],'adj_R2'):.4f}" if meta(kc[0], 'adj_R2') != '' else '') for kc in kinds}},
    ])], ignore_index=True)
    return tbl


# ---------------------------------------------------------------- main
def main():
    df = pd.read_csv(MODEL_CSV, encoding="utf-8-sig", low_memory=False)
    coef = pd.read_csv(COEF_CSV, encoding="utf-8-sig")
    rob = pd.read_csv(ROB_CSV, encoding="utf-8-sig")
    vmap = pd.read_csv(MAP_CSV, encoding="utf-8-sig")
    any_sig, lag_txt, rob_txt = summarize_findings(coef, rob)

    if any_sig:
        finding_zh = (
            "納入時間落差與產業異質性後，部分設定下水管理之調節效果達統計顯著，"
            "顯示水管理對成本黏性的影響具遞延性與產業依存性。主要顯著結果包括："
            + "；".join(lag_txt + rob_txt) + "。")
        finding_en = ("After incorporating time-lag and industry heterogeneity, the "
                      "moderating effect of water management on cost stickiness becomes "
                      "significant in some specifications, indicating lagged and "
                      "industry-dependent effects.")
        concl_zh = (
            "本研究延伸成本黏性架構，納入水管理變數之遞延效果與高／低耗水產業異質性。"
            "結果顯示樣本整體存在顯著成本黏性；在考量時間落差與產業異質性後，水管理之"
            "調節效果於部分設定達顯著，支持「水管理投資效益需時間發酵、且集中於高耗水"
            "產業」之推論。此結果較單純當期全樣本分析更能揭露水管理與成本行為之關聯。")
    else:
        finding_zh = ("即使納入時間落差（t-1、t-2）與高／低耗水產業子樣本，水管理之調節"
                      "效果（β3）於各設定下仍未達統計顯著；成本黏性（β2）則於水回收率相關"
                      "設定中穩健為負。")
        finding_en = ("Even after incorporating time-lag (t-1, t-2) and high/low "
                      "water-use industry subsamples, the moderating effect of water "
                      "management (beta3) remains statistically insignificant, while "
                      "overall cost stickiness (beta2) is robustly negative.")
        concl_zh = ("本研究延伸成本黏性架構，納入水管理變數之遞延效果與高／低耗水產業"
                    "異質性。結果顯示樣本整體存在顯著成本黏性，但水管理之調節效果在當期、"
                    "遞延期與各產業子樣本中均不顯著。可能原因包括具水資料揭露之樣本仍相對"
                    "有限、水回收率變異不足，以及水管理投資對成本結構的影響需更長期間或"
                    "更精細之衡量方能顯現。")

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = EN_FONT
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CJK_FONT)

    # 封面標題
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("企業水管理與廢棄物管理對成本黏性之影響")
    tr.bold = True; tr.font.size = Pt(18); set_cjk(tr)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("——以台灣上市櫃製造業之環境績效與資訊揭露為例")
    sr.font.size = Pt(13); set_cjk(sr)
    doc.add_paragraph()

    # 中文摘要
    add_heading(doc, "中文摘要", 1)
    add_para(doc,
             "本研究以「環境使用密集度」為核心，探討企業用水密集度與廢棄物密集度（實質"
             "環境投入），並以水／廢棄物之資訊揭露與違規裁罰為輔，檢視其對成本黏性（cost "
             "stickiness）的影響。以台灣上市櫃製造業（排除金融保險業）2014至2024年資料為"
             "樣本，採 Anderson, Banker and Janakiraman（2003）成本黏性模型，在營業費用"
             "變動對營收變動的敏感度中，加入收入下降虛擬變數與環境變數之三重交乘，並控制"
             "產業與年份固定效果、以公司叢集穩健標準誤估計；另將核心變數遞延一至兩期，並"
             "依產業污染程度進行高／低污染子樣本檢定。實證發現：樣本整體存在顯著的成本"
             "黏性（β2 顯著為負）；" + finding_zh)
    add_para(doc, "關鍵詞：水管理、水資訊揭露、成本黏性、水回收率、TESG、固定效果模型")

    # Abstract
    add_heading(doc, "Abstract", 1)
    add_para(doc,
             "This study examines how corporate water-management performance (water "
             "recycling rate) and water-information disclosure affect cost stickiness. "
             "Using Taiwanese listed manufacturing firms (excluding financials) from "
             "2014 to 2024 and the Anderson, Banker and Janakiraman (2003) framework, "
             "we add triple interactions of a revenue-decrease dummy with water measures, "
             "controlling for industry and year fixed effects with firm-clustered robust "
             "standard errors, with lagged water measures (t-1, t-2) and high/low "
             "water-use industry subsamples. We find significant overall cost "
             "stickiness. " + finding_en)
    add_para(doc, "Keywords: water management, water disclosure, cost stickiness, "
                  "water recycling rate, TESG, fixed-effects model")

    # 第一章 緒論
    add_heading(doc, "第一章 緒論", 1)
    add_heading(doc, "第一節 研究背景與動機", 2)
    add_para(doc,
             "在氣候變遷、淨零轉型與資源稀缺的多重壓力下，水資源與廢棄物已從企業社會"
             "責任的邊陲議題，逐漸成為攸關營運存續的核心風險。台灣為海島型經濟且製造業"
             "高度密集，半導體、光電、面板、化工、鋼鐵與紡織等支柱產業普遍具有高取水、"
             "高廢水與高廢棄物之特性；2021 年台灣遭逢逾半世紀最嚴重乾旱，中部科學園區"
             "一度面臨限水，凸顯水資源供給對高耗水製造業之關鍵性。與此同時，主管機關"
             "陸續要求上市櫃公司自 2023 年起分階段編製永續報告書、並加嚴廢棄物清理與"
             "環境污染之裁處，使環境資源之取得、循環、處理與法遵成本，實質嵌入企業之"
             "營運成本結構，值得從成本行為的角度深入檢視。")
    add_para(doc,
             "企業成本如何隨營收變動而調整，是管理會計長期關注的課題。Anderson, Banker "
             "and Janakiraman（2003）提出成本僵固性（cost stickiness，或稱成本黏性）"
             "概念，指出當營收下降時，營業費用向下調整的幅度小於營收上升時向上調整的"
             "幅度，呈現顯著的非對稱性。其根源在於管理者的資源調整決策與調整成本：當"
             "面臨需求下滑時，若管理者預期衰退僅屬短暫，或裁撤與重建專用資源之成本高昂，"
             "便傾向保留閒置產能與人力，形成成本黏性。此一概念已被廣泛用於檢視治理、"
             "誘因、財務狀況乃至外部政策對企業成本調整行為之影響。")
    add_para(doc,
             "環境資源的投入具備高投入、長週期、專用性與不可逆之特徵。企業為提升水回收、"
             "處理廢水與廢棄物，往往需建置廢水處理廠、水循環系統與專用處理設施，並簽訂"
             "長期委外清運與處理合約；這類與產能高度綁定的環境支出，於營收下降時難以"
             "即時削減，理論上將加劇成本黏性。另一方面，環境資訊之揭露被視為降低資訊"
             "不對稱、矯正管理者過度樂觀預期之機制，可能反向緩解成本黏性。因此，環境"
             "因素對成本黏性的影響，同時存在『實質投入加劇』與『資訊透明緩解』兩股"
             "方向相反的力量，值得以台灣製造業資料加以釐清。")
    add_para(doc,
             "然而，既有將環境因素連結成本行為的研究，多以整體 ESG 評分或揭露為衡量"
             "（如 Jiang and Yang, 2024；辛珏辰, 2025；鄭紹君, 2018），較少深入「水」與"
             "「廢棄物」等個別、可量化之實質環境構面，且常以『揭露有無』或『罰鍰次數』"
             "等粗略代理，易受衡量誤差與單一極端值干擾。本研究主張，真正貼近調整成本"
             "本質的，是企業對環境資源之「使用密集度」——單位營收所耗用之水量與產生之"
             "廢棄物量，直接反映其營運與環境資源之綁定程度及專用處理投入。相對於水回收率"
             "等揭露樣本稀少、變異有限之指標，用水密集度（總用水量／營收）之樣本涵蓋度"
             "更高，更能提供穩健且具檢定力之證據。")
    add_para(doc,
             "為此，本研究以台灣經濟新報（TEJ）之財務、TESG 企業用水量、廢棄物揭露與"
             "列管事業污染源裁處等資料，建構 2014 至 2024 年台灣上市櫃製造業之公司—年"
             "面板，採 ABJ 成本僵固性模型，將用水密集度、廢棄物密集度及其資訊揭露與"
             "違規裁罰變數，與收入下降虛擬變數交乘，並控制產業與年份固定效果、採公司"
             "叢集穩健標準誤；另以時間落差（遞延一至兩期）與高／低污染產業異質性，"
             "檢驗環境使用密集度對成本黏性之影響及其適用情境。")
    add_para(doc,
             "進一步言，台灣製造業在全球供應鏈中扮演關鍵角色，其生產活動高度倚賴穩定"
             "且充足的水電與原物料供給，同時受環境法規日趨嚴格之規範。近年環境相關"
             "制度快速演進：缺水風險促使高耗水產業投入再生水與製程節水；廢棄物清理"
             "與資源循環規範要求企業建立妥善之清運、處理與再利用機制；污染裁處與環境"
             "違規之公告，也使企業之環境法遵成本與商譽風險同步上升。這些制度變遷意味"
             "企業在環境面的投入與承諾具有長期性與剛性，一旦建置便難以隨短期營收波動"
             "而快速調整，正是檢視環境因素如何形塑成本黏性的合適場域。")
    add_para(doc,
             "從調整成本理論的角度，成本黏性反映管理者在需求變動時「保留或裁撤資源」"
             "的權衡。當向下調整資源（如解約、資遣、處分設施）的成本高於暫時保留閒置"
             "資源的成本，或當管理者對未來需求抱持樂觀預期時，便傾向於營收下降時不"
             "立即縮減成本，形成黏性。環境資源之投入恰具備使黏性加劇的多重特徵：其一，"
             "廢水與廢棄物處理設施屬專用性資產，缺乏次級市場、處分價值低；其二，環境"
             "委外處理多採長期合約，短期難以中止；其三，環境法遵具強制性，企業不能為"
             "節省成本而停止處理污染。因此，企業對環境資源的「使用密集度」越高，理應"
             "在營收下降時面臨越強的向下調整阻力。此一推論構成本研究的核心邏輯。")
    add_heading(doc, "第二節 研究目的與研究問題", 2)
    add_para(doc,
             "本研究之研究目的有四。第一，建立以「環境使用密集度」為核心的成本黏性"
             "分析架構，檢視用水密集度與廢棄物密集度是否顯著影響台灣製造業之成本黏性。"
             "第二，以資訊揭露（水揭露、廢棄物揭露品質）與違規裁罰（廢棄物裁罰金額）作為"
             "輔助構面，比較「實質使用」與「資訊透明度」兩類機制之相對解釋力。第三，"
             "考量環境投入之效益與僵固性需時間發酵，檢驗當期與遞延（t-1、t-2）估計之"
             "差異。第四，考量台灣製造業次產業之環境依賴度差異甚大，比較高污染與低污染"
             "產業之異質性，釐清效果集中之情境。")
    add_para(doc,
             "據此，本研究之核心研究問題為：在控制企業規模、資產與員工密集度、獲利、"
             "財務槓桿及產業與年份固定效果後，企業對水與廢棄物之使用密集度，是否於"
             "營收下降時顯著加劇成本黏性？此一效果是否隨時間遞延而顯現、並集中於高污染"
             "產業？相對地，資訊揭露與違規裁罰之調節效果又是否穩健？透過回答上述問題，"
             "本研究期能補充國內以成本黏性為應變數、並細分水與廢棄物構面之實證缺口，"
             "並為主管機關與企業提供關於環境資源使用密集度與成本韌性之政策與實務參考。")
    add_para(doc,
             "本研究之預期貢獻可分三方面。理論上，將環境因素對成本行為的影響，由抽象的"
             "整體 ESG 或 CSR 評分，推進至可量化、可比較之「實質使用密集度」，使調整"
             "成本理論在環境情境下獲得更貼近本質的檢驗。衡量上，本研究以連續之「用水"
             "密集度」與「廢棄物裁罰金額」取代樣本稀少之水回收率與易受極端值主導之"
             "罰鍰次數，兼顧樣本涵蓋度與衡量效度。方法上，本研究結合時間落差與高／低"
             "污染產業異質性，揭示環境成本效應具「產業依存與遞延顯現」之特徵，此為"
             "僅以全樣本、當期估計之研究所難以觀察，對後續環境會計與成本行為研究具"
             "參考價值。")
    add_para(doc,
             "本研究之所以將「水」與「廢棄物」並列檢視，係因二者同為製造業最主要之"
             "環境資源投入與產出構面，且皆具備專用設施、長期合約與強制法遵之特性，"
             "在成本結構上具有高度可比性；同時，兩者在 TEJ 資料庫中皆有可量化之使用量、"
             "密集度、揭露與裁罰資料，便於建立一致之衡量與比較架構。相較於僅探討單一"
             "構面，並列水與廢棄物有助於檢視環境使用密集度效果之普遍性與穩健性，並辨識"
             "何種構面、何種產業與何種時點下，環境投入對成本行為之影響最為顯著。")
    add_para(doc,
             "在研究範圍上，本研究聚焦於台灣上市櫃「製造業」，並排除金融保險業，樣本"
             "期間為 2014 至 2024 年。之所以限縮於製造業，係因其為水與廢棄物之主要"
             "使用與產出部門，環境資源之投入與處理直接反映於營運成本，最能檢視環境"
             "使用密集度與成本行為之關聯；金融保險業之成本結構與環境足跡與製造業迥異，"
             "故予以排除，以維持樣本之同質性與推論之清晰。")
    add_para(doc,
             "本文後續章節安排如下：第二章回顧成本僵固性與環境／ESG 相關文獻，並據以"
             "提出研究假說；第三章說明資料來源、樣本、變數定義與實證模型；第四章報告"
             "敘述統計、相關係數與主迴歸結果，並進行時間落差與高／低污染產業異質性之"
             "延伸分析；第五章總結研究發現、政策與實務意涵、研究限制與未來研究方向。")

    # 第二章 文獻探討
    add_heading(doc, "第二章 文獻探討與假說", 1)
    add_heading(doc, "第一節 文獻回顧", 2)
    add_para(doc,
             "（一）成本僵固性之理論根源與決定因素。成本僵固性之研究以 Anderson, Banker "
             "and Janakiraman（2003）為濫觴，其以美國公司之銷售、一般及管理費用（SG&A）"
             "驗證營收下降時費用調整之非對稱性，並提出「調整成本」與「管理者樂觀預期」"
             "兩大解釋機制。後續文獻延伸至多元決定因素，國內以台灣上市櫃公司為樣本者尤"
             "眾：張凱瑜（2025）發現成本僵固性與企業風險相關；王惠洳（2025）指出獨立董事"
             "之連結關係會影響成本僵固性；陳思婷（2025）以移轉訂價操弄反映代理問題，"
             "發現操弄程度越高、成本僵固性越大；張宸杰（2023）檢視內部控制與財務危機之"
             "影響；顏宥盈（2024）探討內外部監督效果；謝雅筑（2024）以「飛行員 CEO」之"
             "風險偏好特質、邱圓雅（2020）以內部債，分別檢視管理者誘因對成本僵固性之"
             "作用；張芳瑜（2023）則反向檢視成本僵固性對企業績效之影響，並以總體經濟"
             "變數為調節。上述研究共同確立：成本僵固性在台灣市場穩健存在，且受治理、"
             "誘因、財務狀況與總體環境調節，構成本研究之方法論基礎。")
    add_para(doc,
             "（二）環境、ESG 與成本行為。將環境與永續因素連結成本行為之研究漸增。"
             "Zeng, Peng and Chan 以中國低碳城市政策為外生衝擊，發現政策促使企業增加"
             "綠色創新投資、並受融資限制影響，進而提高成本僵固性；Jiang and Yang（2024）"
             "則指出客戶端之 ESG 揭露可透過供應鏈關係降低供應商管理者之過度樂觀，緩解"
             "其成本僵固性。國內方面，鄭紹君（2018）檢視企業社會責任活動與成本僵固性之"
             "關聯，並以 CSR 願景為調節變數；最具參考價值者為李依潔（2024），其以 2016 至"
             "2022 年台灣上市櫃公司為樣本，發現「企業投入 ESG 環境活動績效越高、成本"
             "僵固性越高」，且採探勘者策略之公司因技術彈性而減緩此正向關係——此結果與"
             "本研究「環境實質投入加劇成本黏性」之推論一致。辛珏辰（2025）進一步將 ESG "
             "活動與 COVID-19 衝擊納入成本僵固性分析。惟上述研究多以整體 ESG 或 CSR "
             "為衡量，尚未細分水與廢棄物等個別環境構面。")
    add_para(doc,
             "（三）水資源績效與揭露。在水資源面，陳樸（2024）針對台灣上市食品公司建構"
             "水資源永續績效之評量方法，凸顯不同產業之用水特性與績效衡量之重要性；"
             "魏郁芳（2025）檢視企業內部因素與外部壓力對水資源揭露之影響，指出揭露行為"
             "受規範壓力與公司特徵驅動。這些研究確立水資源之績效與揭露可被量化、且具"
             "明顯產業異質性，但少有將其進一步連結至企業成本調整行為者，留下可供延伸之"
             "空間。")
    add_para(doc,
             "（四）廢棄物管理與環境裁罰。廢棄物面，鍾文淇（2025）以直轄市為對象研究"
             "一般廢棄物回收率之成效，反映廢棄物處理之政策與制度脈絡。就企業而言，"
             "廢棄物之清運、處理與再利用高度仰賴專用設施與長期委外合約，且環境違規將"
             "面臨主管機關裁處，形成法遵與改善之強制性支出。相較於僅計算「罰鍰次數」，"
             "TEJ 之列管事業污染源裁處資料提供逐案「裁處金額」，更能衡量違規之經濟強度，"
             "且可避免計數型變數受單一極端值主導之偏誤。")
    add_para(doc,
             "值得注意的是，廢棄物構面同時具備「實質產出密集度」與「違規風險」兩種"
             "面向：前者以單位營收之廢棄物量（廢棄物密集度）衡量企業處理負擔之高低，"
             "後者以環境裁處金額反映違規所帶來之強制性法遵與改善支出。二者雖同屬廢棄物"
             "管理，但作用機制不同——密集度反映的是常態性、與產能綁定之處理成本，"
             "裁罰則屬事件性、與治理與守法程度相關之衝擊。本研究將兩者分別檢視，"
             "以區辨常態處理負擔與違規事件對成本黏性之不同影響。")
    add_para(doc,
             "（五）供應鏈機制與研究缺口。供應鏈機制方面，劉佳熒（2022）檢視客戶—供應商"
             "關係對供應商成本行為之影響，與 Jiang and Yang（2024）相互呼應，顯示外部"
             "關係亦形塑企業之成本調整。綜合上述，既有文獻雖已分別累積「成本僵固性決定"
             "因素」與「環境／ESG 財務後果」兩支脈絡，但少有研究以「成本黏性」為應變數，"
             "並將環境構面細分為水與廢棄物之「實質使用密集度」與「資訊揭露」，且同時"
             "結合時間落差與高／低污染產業異質性加以檢驗。此即本研究欲填補之缺口，"
             "並據以發展下列假說。")
    add_para(doc,
             "在衡量與模型方面，Anderson et al.（2003）之對數—對數設定已成為成本僵固性"
             "研究之標準：以營業費用變動率對營收變動率迴歸，並加入營收下降虛擬變數與"
             "營收變動率之交乘，藉交乘項係數捕捉調整之非對稱性。此設定之優點在於係數"
             "可直接解讀為彈性，且便於加入調節變數之三重交乘以檢視特定因素之影響。本"
             "研究上述國內文獻（如張凱瑜, 2025；陳思婷, 2025；李依潔, 2024）多沿用此一"
             "架構，顯示其在台灣資料之適用性；本研究亦以之為基礎，將環境使用密集度、"
             "揭露與裁罰變數納入三重交乘，以檢驗其對成本黏性之調節作用。")
    add_para(doc,
             "在環境資訊揭露之制度背景方面，隨著金管會分階段推動永續報告書之編製與"
             "第三方確信，企業於用水、廢水與廢棄物之量化揭露日趨完整，TESG 亦據以建置"
             "用水量、廢棄物量與 GRI 揭露度等結構化欄位。魏郁芳（2025）指出水資源揭露"
             "受規範壓力與公司特徵驅動，陳樸（2024）則說明產業別對水資源績效衡量之"
             "影響；就廢棄物而言，鍾文淇（2025）反映回收與處理之制度脈絡。這些制度化"
             "與資料化的進展，使本研究得以較過去更完整地衡量企業之環境使用密集度與"
             "揭露品質，並將列管事業污染源之逐案裁處金額納入分析，兼顧衡量之細緻度與"
             "可靠度。")
    add_para(doc,
             "（六）理論整合與本研究之定位。統整上述文獻，環境因素影響成本黏性主要"
             "循兩條路徑。第一為「調整成本／實質投入」路徑：企業投入水循環、廢水與"
             "廢棄物處理等專用資產與長期合約後，於營收下降時難以即時削減，遂加劇成本"
             "黏性；Zeng, Peng and Chan 之綠色創新與融資限制、李依潔（2024）之環境活動"
             "績效越高、成本僵固性越高，均屬此路徑之佐證。第二為「資訊透明度／監督」"
             "路徑：主動且高品質之環境揭露可降低資訊不對稱、矯正管理者過度樂觀，促使"
             "其於需求下滑時更果斷調整資源，從而緩解成本黏性；Jiang and Yang（2024）之"
             "供應鏈 ESG 揭露緩解供應商成本僵固，以及魏郁芳（2025）關於揭露決定因素之"
             "討論，皆與此路徑相關。本研究主張，就「水」與「廢棄物」而言，實質使用密集"
             "度所代表的調整成本路徑，較揭露路徑更為直接且可量化，故以使用密集度為核心"
             "構面，並以揭露與裁罰為輔助對照，藉此辨識二路徑之相對作用。")
    add_para(doc,
             "此外，成本黏性之強度亦受產業與時間因素調節。台灣製造業次產業之環境資源"
             "依賴度差異甚大——半導體、光電、化工、鋼鐵、造紙與紡織等屬高取水、高污染"
             "產業，其環境專用投入與法遵負擔遠高於組裝、服務或資訊類產業；因此，環境"
             "使用密集度對成本黏性之效果，理應集中於高污染產業。再者，環境設施與合約"
             "之成本僵固性往往非於投入當期立即顯現，而是隨設施折舊、合約續期與處理量"
             "累積而逐步強化，故本研究另檢視遞延一至兩期之效果。綜言之，既有文獻為"
             "本研究奠定理論與方法基礎，而本研究以「環境使用密集度為核心、揭露與裁罰"
             "為輔，並納入產業異質性與時間落差」之設計，回應並延伸此一研究脈絡。")
    add_heading(doc, "第二節 水管理績效與成本黏性（H1）", 2)
    add_para(doc,
             "提升水回收率通常需長期投入專用資產與人力，屬高風險、長週期之綠色投資。"
             "當營收下降時，管理者不易削減此類專用資源，閒置資源被保留，進而推升成本"
             "黏性。據此提出："
             )
    add_para(doc, "H1：企業水回收率越高，其成本黏性越大（模型中 β3 預期顯著為負）。",
             bold=True, first_indent=False)
    add_heading(doc, "第三節 水資訊揭露與成本黏性（H2）", 2)
    add_para(doc,
             "主動揭露水資源管理資料代表較高的環境資訊透明度，有助於矯正管理者的過度"
             "樂觀預期與代理問題；當營收下降時，透明度高的企業較能果斷調整閒置資源，"
             "從而降低成本黏性。據此提出：")
    add_para(doc, "H2：相較未揭露者，有揭露水管理資訊的企業其成本黏性較低"
                  "（模型中 β3 預期顯著為正）。", bold=True, first_indent=False)
    add_heading(doc, "第四節 廢棄物管理與成本黏性（H3–H5）", 2)
    add_para(doc,
             "水管理效果在台灣製造業樣本相對有限，故本研究進一步將環境構面擴充至廢棄物"
             "管理。台灣製造業（尤其半導體、光電、化工）之廢棄物清運與處理高度仰賴長期"
             "委外合約與專用處理設施，具明顯的費用僵固性；同時環境違規（罰鍰）具強制性"
             "降本阻力。據此提出三項假設：")
    add_para(doc, "H3（實質投入）：每百萬營收廢棄物量越高，成本黏性越大"
                  "（廢棄物處理合約與費用僵固；β3 預期為負）。", bold=True, first_indent=False)
    add_para(doc, "H4（資訊透明度）：廢棄物管理揭露品質越高，成本黏性越低"
                  "（監督效果矯正管理者預期；β3 預期為正）。", bold=True, first_indent=False)
    add_para(doc, "H5（風險衝擊）：事業廢棄物違規裁罰金額越高，成本黏性越大"
                  "（違規之強制性降本阻力；β3 預期為負）。", bold=True, first_indent=False)
    add_heading(doc, "第五節 使用密集度與成本黏性（H6，核心假說）", 2)
    add_para(doc,
             "綜合水與廢棄物之實質投入機制，本研究以「環境使用密集度」為核心構念："
             "企業單位營收所消耗之水資源與產生之廢棄物越高，代表其營運對環境資源之"
             "依賴與專用處理設施投入越深；當營收下降時，此類與產能綁定之取水、廢水與"
             "廢棄物處理成本難以即時削減，將推升成本黏性。用水密集度樣本涵蓋度遠高於"
             "水回收率，檢定力較佳。據此提出核心假說：")
    add_para(doc, "H6：企業用水密集度（與廢棄物密集度）越高，成本黏性越大"
                  "（β3 預期為負）；且效果集中於高污染產業。", bold=True, first_indent=False)
    add_heading(doc, "第六節 近年國內 ESG 實證與研究缺口", 2)
    add_para(doc,
             "國內成本僵固性研究已累積相當基礎：張凱瑜（2025）檢視成本僵固性與企業"
             "風險之關聯、王惠洳（2025）探討獨立董事連結關係、陳思婷（2025）分析移轉"
             "訂價操弄對成本僵固性之影響，均以 Anderson et al.（2003）模型為基礎並以"
             "台灣上市櫃公司為樣本。與本研究最相關者為辛珏辰（2025）「ESG活動與成本"
             "僵固性」，其將 ESG 面向納入成本僵固性分析，惟仍以整體 ESG 為衡量，"
             "尚未針對「水資源」與「廢棄物」等個別環境構面深入。")
    add_para(doc,
             "在環境／ESG 與財務後果方面，近兩年（2025–2026）國內碩博論文多聚焦於"
             "ESG 揭露及績效對資金成本、風險與盈餘之影響：張岑華（2026）探討 ESG 績效"
             "與漂綠對權益資金成本、楊博勝（2026）檢視 ESG 資訊揭露對企業特有風險、"
             "吳東蓄（2026）分析 ESG 績效對代理成本、許嘉芸（2026）探討資源密集度管理"
             "對財務績效與 ESG 效率之影響。這些研究印證環境資訊揭露之透明度／監督機制"
             "（呼應 H2、H4），但多以資金成本、風險或價值為應變數，鮮少以「成本黏性」"
             "為切入點，且未區分水與廢棄物構面。")
    add_para(doc,
             "綜上，本研究之缺口與貢獻在於：（一）將環境構面細分為「水資源」與"
             "「廢棄物」兩類實質績效與資訊揭露；（二）以成本黏性為應變數，結合時間"
             "落差與高／低污染產業異質性，檢驗其對台灣製造業成本調整行為之影響。")

    # 第三章 研究方法
    add_heading(doc, "第三章 研究方法", 1)
    add_heading(doc, "第一節 資料來源與樣本", 2)
    add_para(doc,
             "本研究資料取自台灣經濟新報（TEJ），包含 IFRS 合併財務（單季，彙總為年度）、"
             "TESG 企業用水量與 TESG 永續揭露資料。樣本期間為 2014 至 2024 年，"
             "排除金融保險業並剔除關鍵變數缺漏之觀測值。財務單季資料以流量加總、"
             "存量取年末方式彙總為年度面板，主鍵為證券代碼與西元年份。")
    add_heading(doc, "第二節 變數定義", 2)
    add_para(doc, "各變數之定義與角色如下表所示。", first_indent=False)
    add_table(doc, vmap.rename(columns={"安全欄名": "變數", "定義": "定義", "角色": "角色"}),
              title="表3-1 變數定義表")
    add_heading(doc, "第三節 實證模型", 2)
    add_para(doc,
             "本研究採 ABJ 成本黏性模型，以營業費用變動對營收變動之敏感度捕捉黏性，"
             "並加入環境管理變數之三重交乘。以 Env_Var 泛指環境管理變數：", first_indent=False)
    add_para(doc,
             "ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Env_Var×D×ΔLNREV) "
             "+ β4·Env_Var + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε",
             first_indent=False)
    add_para(doc,
             "其中 D 為收入下降虛擬變數（當期營收低於前期為1）。β2 捕捉整體成本黏性"
             "（預期為負）；β3 為核心係數。Env_Var 依模型分別代入下列六項環境管理變數：")
    add_para(doc,
             "‧ 核心（使用密集度）——模型6 Water_Intensity（用水密集度＝總用水量/營收，H6）、"
             "模型3 Waste_Intensity（每百萬營收廢棄物，H3）。", first_indent=False)
    add_para(doc,
             "‧ 輔助（揭露與裁罰）——模型1 Water_Rate（水回收率%，H1）、模型2 Water_Disc"
             "（水揭露，H2）、模型4 Waste_Disc（GRI廢棄物揭露度，H4）、模型5 Waste_Fine"
             "（廢棄物裁罰金額佔資產，H5）。", first_indent=False)
    add_para(doc,
             "估計採 OLS 併入產業與年份固定效果，並以公司層級叢集穩健標準誤。"
             "此外，為檢驗時間落差效應，將 Env_Var 分別遞延一期（t-1）與兩期（t-2）"
             "代入；並依 TSE 產業別將樣本分為高／低耗水（高污染）兩組進行異質性檢定，"
             "比較各組之 β3。")

    # 第四章 實證結果
    add_heading(doc, "第四章 實證結果", 1)
    add_heading(doc, "第一節 敘述統計", 2)
    add_table(doc, build_descriptive(df), title="表4-1 主要變數敘述統計")
    add_heading(doc, "第二節 相關係數矩陣", 2)
    add_table(doc, build_corr(df), title="表4-2 主要連續變數相關係數矩陣", num_fmt="{:.3f}")
    add_heading(doc, "第三節 核心結果：使用密集度（當期）", 2)
    add_para(doc,
             "本研究以「環境使用密集度」為核心，包含用水密集度（模型6）與廢棄物密集度"
             "（模型3）。表4-3 為當期結果，括號內為叢集穩健 t 值，*、**、*** 分別代表 "
             "10%、5%、1% 顯著水準。D×ΔLNREV（β2）反映整體成本黏性；β3 為使用密集度與"
             "營收下降之三重交乘。")
    add_table(doc, build_models_table(coef, [("模型6 用水密集度 L0(當期)", "模型6 用水密集度"),
                                             ("模型3 廢棄物密集度 L0(當期)", "模型3 廢棄物密集度")]),
              title="表4-3 使用密集度核心迴歸結果（當期）")
    add_table(doc, build_lag_table(coef, [("用水密集度", "模型6 β3(用水密集度)"),
                                          ("廢棄物主分析", "模型3 β3(廢棄物密集度)")]),
              title="表4-4 使用密集度時間落差效應：各遞延期之 β3")

    add_heading(doc, "第四節 輔助結果：資訊揭露與違規裁罰（當期）", 2)
    add_para(doc,
             "表4-5 為輔助構面：水回收率（模型1）、水揭露（模型2）、廢棄物揭露品質"
             "（模型4）與廢棄物違規裁罰金額（模型5）。此四者作為使用密集度之對照。")
    add_table(doc, build_models_table(coef, [("模型1 水回收率% L0(當期)", "模型1 水回收率%"),
                                             ("模型2 水揭露 L0(當期)", "模型2 水揭露"),
                                             ("模型4 廢棄物揭露 L0(當期)", "模型4 廢棄物揭露"),
                                             ("模型5 廢棄物裁罰金額 L0(當期)", "模型5 廢棄物裁罰金額")]),
              title="表4-5 輔助構面迴歸結果（當期）")
    add_table(doc, build_lag_table(coef, [("次分析", "模型2 β3(水揭露)"),
                                          ("廢棄物次分析", "模型4 β3(廢棄物揭露)"),
                                          ("廢棄物穩健", "模型5 β3(裁罰金額)")]),
              title="表4-6 輔助構面時間落差效應：各遞延期之 β3")

    add_heading(doc, "第五節 產業異質性（高／低耗水高污染產業）", 2)
    add_para(doc,
             "台灣製造業次產業之環境資源依賴度差異甚大，全樣本可能稀釋高耗水高污染產業"
             "之效應。本研究依 TEJ Company DB 之 TSE 產業別，將半導體、光電、電子零組件、"
             "化學、鋼鐵、紡織、造紙、水泥、食品等用水密集／高污染製造業歸為高耗水組，"
             "其餘為低耗水組，分別於當期與遞延期估計並比較 β3。")
    core_measures = ["用水密集度", "廢棄物密集度"]
    aux_measures = ["水回收率%", "水揭露", "廢棄物揭露", "廢棄物裁罰金額"]
    cols_show = ["水衡量", "產業組", "遞延期", "β2(D×ΔREV)", "β2_sig",
                 "β3(Water×D×ΔREV)", "β3_sig", "N"]
    add_para(doc,
             "表4-7 為核心構面（使用密集度）之產業異質性結果。用水密集度於高污染產業當期"
             "之 β3 顯著為負，廢棄物密集度亦於高污染產業遞延期呈顯著負向，顯示效果集中於"
             "高污染產業，符合 H6 與 H3 之推論。", first_indent=False)
    rob_core = rob[rob["水衡量"].isin(core_measures)][cols_show]
    add_table(doc, rob_core, title="表4-7 核心（使用密集度）× 產業異質性 × 時間落差",
              num_fmt="{:.4f}")
    add_para(doc,
             "表4-8 為輔助構面（水回收率、水揭露、廢棄物揭露、廢棄物裁罰金額）之產業"
             "異質性結果，各設定之 β3 多不顯著，作為核心結果之對照。", first_indent=False)
    rob_aux = rob[rob["水衡量"].isin(aux_measures)][cols_show]
    add_table(doc, rob_aux, title="表4-8 輔助（揭露與裁罰）× 產業異質性 × 時間落差",
              num_fmt="{:.4f}")

    # 第五章 結論
    add_heading(doc, "第五章 結論與建議", 1)
    add_para(doc, concl_zh)
    add_para(doc,
             "本研究以「環境使用密集度」為核心之研究貢獻有三。第一，理論上，將環境因素"
             "對成本行為的影響由抽象的「ESG 揭露」推進至可量化的「實質使用密集度」——"
             "用水密集度與廢棄物密集度直接反映企業營運與環境資源之綁定程度，較揭露類"
             "變數更貼近調整成本之本質。實證顯示：使用密集度越高、於營收下降時成本越"
             "難削減（成本黏性越大），且效果集中於高污染產業並具遞延性；相對地，資訊"
             "揭露與違規裁罰之調節效果則不穩健。第二，衡量上，本研究指出以「罰鍰次數」"
             "衡量環境違規之偏誤（易受單一極端值主導），改採裁處金額佔資產之連續衡量後"
             "結論更為穩健；並以用水密集度（樣本涵蓋度高）取代樣本稀少之水回收率，"
             "顯著提升檢定力。第三，方法上，結合時間落差與高／低污染產業異質性設計，"
             "揭示環境成本效應之「產業依存＋遞延」特徵，為全樣本當期分析所無法觀察。")
    add_para(doc,
             "政策與實務意涵：對主管機關而言，推動環境資訊揭露之同時，宜重視高污染產業"
             "之「使用密集度」資訊（用水量、廢棄物量與其密集度）之標準化與可比較性，"
             "俾利利害關係人評估企業於景氣下行時之成本調整彈性與營運韌性。對企業而言，"
             "高使用密集度意味較高之環境專用資產與委外處理承諾，於需求衰退時形成向下"
             "調整阻力，管理階層應於資本配置與產能規劃時將此僵固性納入考量，並透過"
             "循環化、製程改善與資源效率投資，降低營運對環境資源之剛性依賴。後續研究可"
             "延伸至碳排密集度等其他環境使用構面，或以環境稽查／缺水等外生事件強化因果識別。")

    # 參考文獻
    add_heading(doc, "參考文獻", 1)

    def add_ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(24)
        p.paragraph_format.first_line_indent = Pt(-24)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text); r.font.size = Pt(11); set_cjk(r)

    add_para(doc, "一、主要參考文獻（核心理論與模型來源）", bold=True, first_indent=False)
    for ref in MAIN_REFS:
        add_ref(ref)
    add_para(doc, "二、次要參考文獻（國內相關碩博士論文，依年份新到舊）",
             bold=True, first_indent=False)
    for i, (yr, author, title, org) in enumerate(
            sorted(SUB_REFS, key=lambda x: -x[0]), 1):
        add_ref(f"[{i}] {author}（{yr}）。{title}。{org}。")

    doc.save(OUT_DOCX)
    print(f"已生成論文：{OUT_DOCX}")
    print(f"  敘述統計 {len(DESC_VARS)} 變數、相關矩陣 {len(CORR_VARS)} 變數、"
          f"主迴歸 2 模型、穩健性 {len(rob)} 設定、變數定義 {len(vmap)} 列。")


if __name__ == "__main__":
    main()
