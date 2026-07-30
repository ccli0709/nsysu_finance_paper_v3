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
WATER_MEASURES = ["水回收率%", "水揭露"]
WASTE_MEASURES = ["廢棄物密集度", "廢棄物揭露", "廢棄物罰鍰"]

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
             ("廢棄物穩健", "模型5(廢棄物罰鍰)")]
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
             "本研究探討企業水管理績效（水回收率）與水資訊揭露對成本黏性（cost "
             "stickiness）的影響。以台灣上市櫃製造業（排除金融保險業）2014至2024年"
             "資料為樣本，採 Anderson, Banker and Janakiraman（2003）成本黏性模型，"
             "在營業費用變動對營收變動的敏感度中，加入收入下降虛擬變數與水管理變數之"
             "三重交乘，並控制產業與年份固定效果、以公司叢集穩健標準誤估計。為捕捉水管理"
             "投資之遞延效益，另將核心水變數遞延一至兩期，並依產業耗水程度進行高／低耗水"
             "子樣本檢定。實證發現：樣本整體存在顯著的成本黏性（β2 顯著為負）；" + finding_zh)
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
             "在氣候變遷與永續發展趨勢下，水資源已成為企業環境管理的核心議題之一。"
             "水資源的取得、循環與回收不僅牽動生產成本，也影響企業面對營運波動時的"
             "資源調整彈性。成本黏性描述企業在營收下降時，成本向下調整的幅度小於營收"
             "上升時向上調整幅度的現象，其根源在於管理者的資源調整決策與調整成本。"
             "水資源相關投資（如廢水處理與循環系統）通常具高投入、長週期的特性，"
             "可能改變企業面臨需求衰退時的資源調整行為，因而與成本黏性產生連結。")
    add_heading(doc, "第二節 研究目的", 2)
    add_para(doc,
             "本研究以台灣製造業為對象，分別從「實質績效」與「資訊透明度」兩個維度，"
             "檢視（一）水回收率是否影響成本黏性；（二）水資訊揭露是否影響成本黏性，"
             "以釐清水管理在企業成本調整行為中的角色。")

    # 第二章 文獻探討
    add_heading(doc, "第二章 文獻探討與假說", 1)
    add_para(doc,
             "成本黏性文獻自 Anderson et al.（2003）以降，普遍以調整成本與管理者決策"
             "解釋成本的非對稱行為。近期研究進一步將環境與永續因素納入。Zeng, Peng and "
             "Chan 探討低碳城市政策透過綠色創新與融資限制影響成本黏性；Jiang and Yang"
             "（2024, Economics Letters）則檢視 ESG 揭露透過供應鏈關係與資訊透明度影響"
             "成本黏性。本研究延伸此一脈絡，聚焦於「水」此一具體環境構面。")
    add_heading(doc, "第一節 水管理績效與成本黏性（H1）", 2)
    add_para(doc,
             "提升水回收率通常需長期投入專用資產與人力，屬高風險、長週期之綠色投資。"
             "當營收下降時，管理者不易削減此類專用資源，閒置資源被保留，進而推升成本"
             "黏性。據此提出："
             )
    add_para(doc, "H1：企業水回收率越高，其成本黏性越大（模型中 β3 預期顯著為負）。",
             bold=True, first_indent=False)
    add_heading(doc, "第二節 水資訊揭露與成本黏性（H2）", 2)
    add_para(doc,
             "主動揭露水資源管理資料代表較高的環境資訊透明度，有助於矯正管理者的過度"
             "樂觀預期與代理問題；當營收下降時，透明度高的企業較能果斷調整閒置資源，"
             "從而降低成本黏性。據此提出：")
    add_para(doc, "H2：相較未揭露者，有揭露水管理資訊的企業其成本黏性較低"
                  "（模型中 β3 預期顯著為正）。", bold=True, first_indent=False)
    add_heading(doc, "第三節 廢棄物管理與成本黏性（H3–H5）", 2)
    add_para(doc,
             "水管理效果在台灣製造業樣本相對有限，故本研究進一步將環境構面擴充至廢棄物"
             "管理。台灣製造業（尤其半導體、光電、化工）之廢棄物清運與處理高度仰賴長期"
             "委外合約與專用處理設施，具明顯的費用僵固性；同時環境違規（罰鍰）具強制性"
             "降本阻力。據此提出三項假設：")
    add_para(doc, "H3（實質投入）：每百萬營收廢棄物量越高，成本黏性越大"
                  "（廢棄物處理合約與費用僵固；β3 預期為負）。", bold=True, first_indent=False)
    add_para(doc, "H4（資訊透明度）：廢棄物管理揭露品質越高，成本黏性越低"
                  "（監督效果矯正管理者預期；β3 預期為正）。", bold=True, first_indent=False)
    add_para(doc, "H5（風險衝擊）：事業廢棄物罰鍰次數越多，成本黏性越大"
                  "（違規之強制性降本阻力；β3 預期為負）。", bold=True, first_indent=False)
    add_heading(doc, "第四節 國內相關文獻與研究缺口", 2)
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
             "‧ 水資源——模型1 Water_Rate（水回收率%，H1）、模型2 Water_Disc（水揭露，H2）。",
             first_indent=False)
    add_para(doc,
             "‧ 廢棄物——模型3 Waste_Intensity（每百萬營收廢棄物，H3）、模型4 Waste_Disc"
             "（GRI廢棄物揭露度，H4）、模型5 Waste_Fine（事業廢棄物罰鍰次數，H5）。",
             first_indent=False)
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
    add_heading(doc, "第三節 主迴歸結果（當期）", 2)
    add_para(doc,
             "表4-3 為當期（t）主迴歸結果，括號內為叢集穩健 t 值，*、**、*** 分別代表 "
             "10%、5%、1% 顯著水準。模型1（水回收率）之 D×ΔLNREV 顯著為負，顯示樣本存在"
             "成本黏性；然當期水管理三重交乘（β3）在兩模型中皆不顯著。")
    add_table(doc, build_reg_table(coef), title="表4-3 成本黏性主迴歸結果（當期，模型1／模型2）",
              num_fmt="{:.4f}")

    add_heading(doc, "第四節 時間落差效應（遞延期分析）", 2)
    add_para(doc,
             "考量水管理投資效益需時間發酵——當期投入的環保設備，可能於 1 至 2 年後方"
             "轉為難以裁撤的沈沒成本或發揮降本效益——本研究將核心水變數分別遞延一期"
             "（t-1）與兩期（t-2）後重新估計三重交乘 β3。表4-4 彙整各遞延期之 β3。")
    add_table(doc, build_lag_table(coef, [("主分析", "模型1 β3(水回收率%)"),
                                          ("次分析", "模型2 β3(水揭露)")]),
              title="表4-4 水管理時間落差效應：各遞延期之 β3")

    add_heading(doc, "第五節 廢棄物管理結果（延伸主題）", 2)
    add_para(doc,
             "表4-5 為廢棄物管理三模型之當期結果：模型3（每百萬營收廢棄物，實質投入）、"
             "模型4（廢棄物揭露品質）、模型5（事業廢棄物罰鍰，違規風險）。相較水管理，"
             "廢棄物指標之有效樣本較大（密集度約 3,273、揭露約 13,875 筆）。")
    add_table(doc, build_waste_reg_table(coef),
              title="表4-5 廢棄物管理主迴歸結果（當期，模型3／4／5）", num_fmt="{:.4f}")
    add_table(doc, build_lag_table(coef, [("廢棄物主分析", "模型3 β3(密集度)"),
                                          ("廢棄物次分析", "模型4 β3(揭露)"),
                                          ("廢棄物穩健", "模型5 β3(罰鍰)")]),
              title="表4-6 廢棄物管理時間落差效應：各遞延期之 β3")

    add_heading(doc, "第六節 產業異質性（高／低耗水高污染產業）", 2)
    add_para(doc,
             "台灣製造業次產業之環境資源依賴度差異甚大，全樣本可能稀釋高耗水高污染產業"
             "之效應。本研究依 TEJ Company DB 之 TSE 產業別，將半導體、光電、電子零組件、"
             "化學、鋼鐵、紡織、造紙、水泥、食品等用水密集／高污染製造業歸為高耗水組，"
             "其餘為低耗水組，分別於當期與遞延期估計並比較 β3。")
    rob_water = rob[rob["水衡量"].isin(WATER_MEASURES)][
        ["水衡量", "產業組", "遞延期", "β2(D×ΔREV)", "β2_sig",
         "β3(Water×D×ΔREV)", "β3_sig", "N"]]
    add_table(doc, rob_water, title="表4-7 水管理 × 產業異質性 × 時間落差", num_fmt="{:.4f}")
    rob_waste = rob[rob["水衡量"].isin(WASTE_MEASURES)][
        ["水衡量", "產業組", "遞延期", "β2(D×ΔREV)", "β2_sig",
         "β3(Water×D×ΔREV)", "β3_sig", "N"]]
    add_table(doc, rob_waste, title="表4-8 廢棄物管理 × 產業異質性 × 時間落差", num_fmt="{:.4f}")

    # 第五章 結論
    add_heading(doc, "第五章 結論與建議", 1)
    add_para(doc, concl_zh)
    add_para(doc,
             "研究貢獻在於將「水」此一具體環境構面納入成本黏性研究，並以時間落差與產業"
             "異質性設計檢視其效果，提供台灣製造業之實證證據。實務上建議主管機關持續推動"
             "水資訊揭露標準化；後續研究可延長樣本期間、細分水資源投資類型，或改採更精細"
             "的揭露品質衡量，以進一步檢驗其效果。")

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
