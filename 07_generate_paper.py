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
OUT_DOCX = "水管理與成本黏性_論文.docx"

CJK_FONT = "新細明體"
EN_FONT = "Times New Roman"

DESC_VARS = ["Y_dLNSGA", "dLNREV", "D", "Water_Rate", "Water_Disc",
             "Size", "AI", "EI", "ROA", "Lev", "Decrease"]
CORR_VARS = ["Y_dLNSGA", "dLNREV", "Water_Rate", "Size", "AI", "EI", "ROA", "Lev"]


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
    """模型1/2 併排：變數 | 模型1 係數(顯著) | 模型2 係數(顯著)。"""
    labels = coef["模型"].unique().tolist()
    m1 = labels[0]
    m2 = labels[1] if len(labels) > 1 else labels[0]
    order = ["dLNREV", "D_x_dREV", "WaterRate_D_dREV", "WaterDisc_D_dREV",
             "Water_Rate", "Water_Disc", "Size_D_dREV", "AI_D_dREV",
             "EI_D_dREV", "ROA_D_dREV", "Lev_D_dREV", "Decrease_D_dREV"]

    def cell(label, var):
        r = coef[(coef["模型"] == label) & (coef["變數"] == var)]
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
    # N / adjR2 附註列
    def meta(label, key):
        r = coef[coef["模型"] == label]
        return "" if r.empty else r[key].iloc[0]
    n1, n2 = meta(m1, "N"), meta(m2, "N")
    a1, a2 = meta(m1, "adj_R2"), meta(m2, "adj_R2")
    tbl = pd.concat([tbl, pd.DataFrame([
        {"變數": "N", "模型1(水回收率%)": f"{int(n1):,}", "模型2(水揭露)": f"{int(n2):,}"},
        {"變數": "adj. R²", "模型1(水回收率%)": f"{a1:.4f}", "模型2(水揭露)": f"{a2:.4f}"},
    ])], ignore_index=True)
    return tbl


# ---------------------------------------------------------------- main
def main():
    df = pd.read_csv(MODEL_CSV, encoding="utf-8-sig", low_memory=False)
    coef = pd.read_csv(COEF_CSV, encoding="utf-8-sig")
    rob = pd.read_csv(ROB_CSV, encoding="utf-8-sig")
    vmap = pd.read_csv(MAP_CSV, encoding="utf-8-sig")

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = EN_FONT
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CJK_FONT)

    # 封面標題
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("企業水管理績效與水資訊揭露對成本黏性之影響")
    tr.bold = True; tr.font.size = Pt(18); set_cjk(tr)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("——以台灣上市櫃製造業為例")
    sr.font.size = Pt(13); set_cjk(sr)
    doc.add_paragraph()

    # 中文摘要
    add_heading(doc, "中文摘要", 1)
    add_para(doc,
             "本研究探討企業水管理績效（水回收率）與水資訊揭露對成本黏性（cost "
             "stickiness）的影響。以台灣上市櫃製造業（排除金融保險業）2014至2024年"
             "資料為樣本，採 Anderson, Banker and Janakiraman（2003）成本黏性模型，"
             "在營業費用變動對營收變動的敏感度中，加入收入下降虛擬變數與水管理變數之"
             "三重交乘，並控制產業與年份固定效果、以公司叢集穩健標準誤估計。實證發現："
             "樣本整體存在顯著的成本黏性（β2 顯著為負）；惟水回收率與水資訊揭露對成本"
             "黏性之調節效果在主分析與各項穩健性檢定中均未達統計顯著。結果顯示，在現階段"
             "台灣製造業樣本中，水管理之實質績效與揭露尚未對成本調整行為產生可量測的影響。")
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
             "standard errors. We find significant overall cost stickiness, but neither "
             "the water recycling rate nor water disclosure significantly moderates cost "
             "stickiness across the main and robustness specifications.")
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
             "並加入水管理變數之三重交乘：", first_indent=False)
    add_para(doc,
             "ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Water_Var×D×ΔLNREV) "
             "+ β4·Water_Var + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε",
             first_indent=False)
    add_para(doc,
             "其中 D 為收入下降虛擬變數（當期營收低於前期為1）。β2 捕捉整體成本黏性"
             "（預期為負）；β3 為核心係數：主分析以水回收率代入，次分析以水揭露虛擬"
             "變數代入。估計採 OLS 併入產業與年份固定效果，並以公司層級叢集穩健標準誤。")

    # 第四章 實證結果
    add_heading(doc, "第四章 實證結果", 1)
    add_heading(doc, "第一節 敘述統計", 2)
    add_table(doc, build_descriptive(df), title="表4-1 主要變數敘述統計")
    add_heading(doc, "第二節 相關係數矩陣", 2)
    add_table(doc, build_corr(df), title="表4-2 主要連續變數相關係數矩陣", num_fmt="{:.3f}")
    add_heading(doc, "第三節 主迴歸結果", 2)
    add_para(doc,
             "表4-3 為主迴歸結果，括號內為叢集穩健 t 值，*、**、*** 分別代表 10%、5%、"
             "1% 顯著水準。模型1（水回收率）之 D×ΔLNREV 顯著為負，顯示樣本存在成本黏性；"
             "然水管理三重交乘（β3）在兩模型中皆不顯著，H1 與 H2 未獲支持。")
    add_table(doc, build_reg_table(coef), title="表4-3 成本黏性主迴歸結果（模型1／模型2）",
              num_fmt="{:.4f}")
    add_heading(doc, "第四節 穩健性檢定", 2)
    add_para(doc,
             "為檢驗結論穩健性，分別替換水績效衡量（製程水回收率%）、水揭露定義"
             "（GRI 揭露度）、營業費用定義（推銷＋管理費用）與產業樣本（限水資料充足產業）。"
             "如表4-4，成本黏性（β2）於水回收率相關設定中穩健為負，惟核心調節係數（β3）"
             "於所有設定下均不顯著，與主結果一致。")
    rob_show = rob[["設定", "β2(D×ΔREV)", "β2_sig", "β3(Water×D×ΔREV)", "β3_sig", "N"]]
    add_table(doc, rob_show, title="表4-4 穩健性檢定跨設定對照", num_fmt="{:.4f}")

    # 第五章 結論
    add_heading(doc, "第五章 結論與建議", 1)
    add_para(doc,
             "本研究以台灣製造業 2014–2024 年資料，檢視水管理績效與水資訊揭露對成本黏性"
             "之影響。實證結果顯示樣本整體存在顯著成本黏性，但水回收率與水資訊揭露之"
             "調節效果在主分析與穩健性檢定中均不顯著。可能原因包括：具水回收率揭露之"
             "樣本仍相對有限、水資訊揭露的品質與可比較性尚待提升，以及水管理投資對成本"
             "結構的影響需更長期間方能顯現。")
    add_para(doc,
             "研究貢獻在於將「水」此一具體環境構面納入成本黏性研究，並提供台灣製造業的"
             "初步證據。實務上建議主管機關持續推動水資訊揭露標準化；後續研究可延長樣本"
             "期間、細分水資源投資類型，或改採更精細的揭露品質衡量，以進一步檢驗其效果。")

    # 參考文獻
    add_heading(doc, "參考文獻", 1)
    for ref in [
        "Anderson, M. C., Banker, R. D., & Janakiraman, S. N. (2003). Are selling, "
        "general, and administrative costs \u201csticky\u201d? Journal of Accounting "
        "Research, 41(1), 47\u201363.",
        "Jiang, W., & Yang, W. (2024). ESG disclosure and corporate cost stickiness: "
        "Evidence from supply-chain relationships. Economics Letters, 238, 111697.",
        "Zeng, J., Peng, M., & Chan, K. C. The impact of low-carbon city policy on "
        "corporate cost stickiness. (working paper).",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(24)
        p.paragraph_format.first_line_indent = Pt(-24)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(ref); r.font.size = Pt(11); set_cjk(r)

    doc.save(OUT_DOCX)
    print(f"已生成論文：{OUT_DOCX}")
    print(f"  敘述統計 {len(DESC_VARS)} 變數、相關矩陣 {len(CORR_VARS)} 變數、"
          f"主迴歸 2 模型、穩健性 {len(rob)} 設定、變數定義 {len(vmap)} 列。")


if __name__ == "__main__":
    main()
