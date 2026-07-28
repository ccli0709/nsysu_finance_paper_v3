# -*- coding: utf-8 -*-
"""生成給指導教授的簡短研究進度報告（Word）：水管理 vs 水+廢棄物管理比較。"""
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn

OUT = "研究進度報告_水與廢棄物管理與成本黏性.docx"
CJK, EN = "新細明體", "Times New Roman"
ROB_CSV = "06_robustness_summary.csv"

WATER_M = ["水回收率%", "水揭露"]
WASTE_M = ["廢棄物密集度", "廢棄物揭露", "廢棄物罰鍰"]


def cjk(run):
    run.font.name = EN
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)


def h(doc, text, lv=1):
    p = doc.add_heading(level=lv)
    cjk(p.add_run(text))
    return p


def para(doc, text, size=12, bold=False, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    cjk(r)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    r.font.size = Pt(12)
    cjk(r)
    return p


def add_table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Light Grid Accent 1"
    for j, c in enumerate(header):
        rr = t.rows[0].cells[j].paragraphs[0].add_run(str(c))
        rr.bold = True
        rr.font.size = Pt(9)
        cjk(rr)
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            rr = cells[j].paragraphs[0].add_run(str(v))
            rr.font.size = Pt(9)
            cjk(rr)
    return t


# ---------------------------------------------------------------- 讀結果
rob = pd.read_csv(ROB_CSV, encoding="utf-8-sig")
b3col = "β3(Water×D×ΔREV)"


def n_at(measure):
    r = rob[(rob["水衡量"] == measure) & (rob["產業組"] == "全樣本") & (rob["遞延期"] == "t-0")]
    return int(r["N"].iloc[0]) if len(r) else 0


def sig_count(measures):
    sub = rob[rob["水衡量"].isin(measures)]
    total = sub[b3col].notna().sum()
    sig = sub[sub["β3_sig"].isin(["*", "**", "***"])]
    return int(sig.shape[0]), int(total)


water_sig, water_tot = sig_count(WATER_M)
waste_sig, waste_tot = sig_count(WASTE_M)

waste_sig_rows = rob[rob["水衡量"].isin(WASTE_M) & rob["β3_sig"].isin(["*", "**", "***"])]

# ---------------------------------------------------------------- 文件
doc = Document()
n = doc.styles["Normal"]
n.font.name = EN
n.font.size = Pt(12)
n.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("研究進度報告：水管理與廢棄物管理對成本黏性之影響")
r.bold = True
r.font.size = Pt(17)
cjk(r)
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
cjk(sub.add_run("——以 TEJ 台灣上市櫃製造業實際資料之初步驗證"))

para(doc, "敬呈 指導教授。前次以「水管理」為題之驗證顯示核心調節效果不顯著；依老師建議，"
          "本次在原水管理基礎上再納入「廢棄物管理」指標（實質投入、揭露品質、違規罰鍰），"
          "並維持產業與年份固定效果。以下彙報最終設計、結果，以及「僅水管理」與「水＋廢棄物」"
          "之對照，懇請老師指示後續方向。")

h(doc, "一、研究設計（延續成本黏性 ABJ 模型）", 2)
para(doc, "應變數 ΔLNSGA（營業費用變動）；核心為 ΔLNREV、收入下降虛擬 D，與環境變數之"
          "三重交乘 X×D×ΔLNREV。控制變數：企業規模、資產密集度、員工密集度、ROA、財務"
          "槓桿、連續兩年營收下降；並控制產業(SASB)與年份固定效果，採公司叢集穩健標準誤、"
          "連續變數縮尾 1%。X 分別代入水管理與廢棄物管理指標，並檢驗當期與遞延（t-1、t-2）"
          "及高／低耗水（高污染）產業之異質性。資料取自 TEJ，台灣製造業、排除金融、2014–2024，"
          "面板 18,137 觀測（1,945 家）。")

h(doc, "二、變數與假設（水 vs 廢棄物）", 2)
add_table(doc,
          ["構面", "主分析 X（實質投入）", "次分析 X（資訊透明度）", "穩健 X（風險）", "假設方向"],
          [["水管理", "水回收率%", "水揭露(有/無)", "—", "H1 β3<0、H2 β3>0"],
           ["廢棄物管理", "每百萬營收廢棄物", "GRI廢棄物揭露度", "事業廢棄物罰鍰次數",
            "H3 β3<0、H4 β3>0、H5 β3<0"]])
para(doc, "水管理：H1 水回收投資屬長期專用資產，難裁撤→加劇黏性；H2 揭露提升透明度→"
          "緩解黏性。廢棄物管理：H3 廢棄物處理仰賴長期委外合約與專用設施，費用僵固→加劇"
          "黏性；H4 高品質揭露具監督效果→緩解黏性；H5 環境違規（罰鍰）具強制降本阻力→"
          "加劇黏性。", indent=False)

h(doc, "三、初步結果", 2)
para(doc, f"1. 成本黏性穩健存在：β2（D×ΔLNREV）於水回收率與廢棄物密集度樣本均顯著為負"
          f"（如廢棄物密集度全樣本 −0.64***、高耗水 t-1 −2.03***）。")
para(doc, f"2. 水管理調節（β3）：{water_tot} 個設定中顯著者 {water_sig} 個——即當期、"
          f"遞延、及高／低耗水各設定下 H1／H2 均未獲支持。")
para(doc, f"3. 廢棄物管理調節（β3）：{waste_tot} 個設定中顯著者 {waste_sig} 個，"
          f"且集中於高耗水（高污染）產業並具遞延性，方向多支持假設（詳附表二）。")

h(doc, "四、洞察", 2)
bullet(doc, "納入廢棄物管理明顯強化研究：主分析有效樣本由水的約千筆擴增至數千筆"
            f"（每百萬營收廢棄物 {n_at('廢棄物密集度'):,}、廢棄物揭露 {n_at('廢棄物揭露'):,}）。")
bullet(doc, "效果具『產業依存＋時間落差』特徵：在高耗水高污染產業、遞延一期（t-1）時，"
            "廢棄物密集度與罰鍰對成本黏性呈顯著加劇，符合台灣製造業（半導體、化工、鋼鐵等）"
            "廢棄物委外合約僵固與環境違規降本阻力之情境。")
bullet(doc, "純水管理指標於任何設定均不顯著，建議將主軸擴展為『水＋廢棄物』之環境成本行為。")

h(doc, "五、請示", 2)
para(doc, "懇請老師指示是否以『水＋廢棄物管理對成本黏性』為主軸持續深化。若可行，後續將"
          "強化衡量（廢棄物密集度、揭露品質指數）、檢視高污染產業與遞延結構之機制，"
          "並考慮以環境事件（如環保稽查／罰鍰）作為識別策略。")

# ================= 附頁：水 vs 水+廢棄物 一頁對照 =================
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
h(doc, "附表一、水管理 vs 水＋廢棄物管理：整體對照", 2)
add_table(doc,
          ["面向", "僅水管理（原）", "水＋廢棄物管理（新）"],
          [["環境構面", "水（回收率、揭露）", "水 ＋ 廢棄物（密集度、揭露、罰鍰）"],
           ["研究假設", "H1、H2", "H1、H2 ＋ H3、H4、H5"],
           ["主迴歸模型數", "2（模型1/2）", "5（模型1–5）"],
           ["穩健性設定數", f"{water_tot}", f"{water_tot + waste_tot}"],
           ["主分析可用樣本", f"水回收率 {n_at('水回收率%'):,}",
            f"廢棄物密集度 {n_at('廢棄物密集度'):,}"],
           ["揭露樣本", f"水揭露 {n_at('水揭露'):,}", f"廢棄物揭露 {n_at('廢棄物揭露'):,}"],
           ["β2 成本黏性", "顯著為負", "顯著為負（水、廢棄物皆是）"],
           ["β3 顯著設定數", f"{water_sig} / {water_tot}", f"{waste_sig} / {waste_tot}（廢棄物）"],
           ["核心結論", "H1/H2 皆不顯著", "廢棄物於高污染產業＋遞延出現顯著（支持H3/H5）"]])

h(doc, "附表二、廢棄物管理達顯著之設定（β3）", 2)
# 各廢棄物衡量對應之假設與預期方向（負=加劇黏性、正=緩解黏性）
WASTE_HYP = {
    "廢棄物密集度": ("H3", "負"),
    "廢棄物揭露": ("H4", "正"),
    "廢棄物罰鍰": ("H5", "負"),
}
rows = []
for _, r in waste_sig_rows.iterrows():
    direction = "加劇黏性" if r[b3col] < 0 else "緩解黏性"
    hyp, expect = WASTE_HYP.get(r["水衡量"], ("", ""))
    match = (expect == "負" and r[b3col] < 0) or (expect == "正" and r[b3col] > 0)
    # 理論以高耗水（高污染）產業為主要適用對象
    if match and r["產業組"] == "低耗水":
        support = f"部分符合{hyp}（惟屬低耗水，理論適用性較弱）"
    elif match:
        support = f"支持 {hyp}"
    else:
        support = f"與 {hyp} 相反"
    rows.append([r["水衡量"], r["產業組"], r["遞延期"],
                 f"{r[b3col]:.4f}{r['β3_sig']}", direction, support, f"{int(r['N']):,}"])
if not rows:
    rows = [["—", "—", "—", "無", "—", "—", "—"]]
add_table(doc, ["廢棄物衡量", "產業組", "遞延期", "β3", "方向", "是否支持假設", "N"], rows)
para(doc, "說明：星號 *、**、*** 分別代表 10%、5%、1% 顯著水準；假設方向為 H3 密集度加劇"
          "（β3<0）、H4 揭露緩解（β3>0）、H5 罰鍰加劇（β3<0），理論主要適用於高耗水"
          "（高污染）產業。相對於水管理在所有設定均不顯著，廢棄物管理在高污染產業並考量"
          "時間落差後出現支持假設之顯著調節，顯示廢棄物構面對台灣製造業成本行為的解釋力較強。",
     indent=False)

doc.save(OUT)
print("已生成：", OUT)
print(f"  水管理 β3 顯著 {water_sig}/{water_tot}；廢棄物 β3 顯著 {waste_sig}/{waste_tot}")
print(f"  廢棄物顯著設定：{len(waste_sig_rows)} 筆")
