# -*- coding: utf-8 -*-
"""生成給指導教授的簡短研究進度報告（Word）。"""
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn

OUT = "研究進度報告_水管理與成本黏性.docx"
CJK, EN = "新細明體", "Times New Roman"


def cjk(run):
    run.font.name = EN
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)


def h(doc, text, lv=1):
    p = doc.add_heading(level=lv)
    cjk(p.add_run(text))
    return p


def para(doc, text, size=12, bold=False, indent=True, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
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


def add_summary_table(doc, csv_path):
    df = pd.read_csv(csv_path, encoding="utf-8-sig")

    def fmt(row):
        b2 = f"{row['β2(D×ΔREV)']:.4f}"
        b2s = "" if row["β2_sig"] == "n.s." else row["β2_sig"]
        b3 = f"{row['β3(Water×D×ΔREV)']:.4f}"
        b3s = "" if row["β3_sig"] == "n.s." else row["β3_sig"]
        return b2 + b2s, b3 + b3s

    cols = ["水衡量", "產業組", "遞延期", "β2 (成本黏性)", "β3 (水管理調節)", "N"]
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = "Light Grid Accent 1"
    for j, c in enumerate(cols):
        rr = t.rows[0].cells[j].paragraphs[0].add_run(c)
        rr.bold = True
        rr.font.size = Pt(9)
        cjk(rr)
    for _, row in df.iterrows():
        b2, b3 = fmt(row)
        vals = [row["水衡量"], row["產業組"], row["遞延期"], b2, b3, f"{int(row['N']):,}"]
        cells = t.add_row().cells
        for j, v in enumerate(vals):
            rr = cells[j].paragraphs[0].add_run(str(v))
            rr.font.size = Pt(9)
            cjk(rr)


doc = Document()
n = doc.styles["Normal"]
n.font.name = EN
n.font.size = Pt(12)
n.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("研究進度報告：企業水管理與成本黏性")
r.bold = True
r.font.size = Pt(17)
cjk(r)
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
cjk(sub.add_run("——以 TEJ 台灣上市櫃製造業實際資料之初步驗證"))

para(doc, "敬呈 指導教授。以下就老師建議之研究方向，說明我以 TEJ 實際資料所做的"
          "初步驗證、最終變數設計與假設、以 AI 輔助跑出的初步結果與洞察，"
          "懇請老師指示此方向是否可行，以作為後續深化之依據。", bold=False)

h(doc, "一、研究方向（老師建議）", 2)
para(doc, "以「成本黏性（cost stickiness）」為核心，參酌兩篇文獻——Zeng, Peng and "
          "Chan（低碳城市政策、綠色創新）與 Jiang and Yang（2024, Economics "
          "Letters；ESG 揭露）——探討「企業水管理績效」與「水資訊揭露」對成本黏性的"
          "影響，並依老師提示納入「時間落差效應」與「產業異質性」兩項延伸設計。")

h(doc, "二、資料與樣本", 2)
para(doc, "資料全數取自 TEJ，包含：IFRS 合併財務（單季，彙總為年度）、TESG 企業用水量"
          "（年度）、TESG 永續揭露，以及 TEJ Company DB 產業別（TSE 產業）。以台灣"
          "上市櫃製造業為對象、排除金融保險業、期間 2014–2024，建構「證券代碼×年度」"
          "面板。清理後樣本 18,137 個觀測（1,945 家）；迴歸可用 16,661 筆；其中具"
          "水回收率者約 996 筆。")

h(doc, "三、最終變數設計", 2)
bullet(doc, "應變數：ΔLNSGA＝營業費用變動（取對數差分）。")
bullet(doc, "核心自變數：ΔLNREV（營收變動）、D（營收下降虛擬）；三重交乘 "
            "Water_Var × D × ΔLNREV。")
bullet(doc, "水管理衡量：Water_Rate＝水回收率%（績效，主分析）；Water_Disc＝是否揭露"
            "水資料（透明度，次分析）。")
bullet(doc, "控制變數：企業規模、資產密集度、員工密集度、ROA、財務槓桿、連續兩年"
            "營收下降；並控制產業與年份固定效果，採公司叢集穩健標準誤，連續變數縮尾 1%。")
bullet(doc, "延伸設計：核心水變數遞延一至兩期（t-1、t-2）；依 TSE 產業別精準界定"
            "高耗水產業（半導體、光電、電子零組件、紡織、造紙、鋼鐵、化學、水泥、"
            "食品、塑膠等）與低耗水產業進行子樣本比較。")

h(doc, "四、研究假設", 2)
para(doc, "H1：水回收率越高，成本黏性越大（因水循環設備屬長期專用投資，難以裁撤；"
          "預期 β3 顯著為負）。", indent=False)
para(doc, "H2：有揭露水管理資訊者，成本黏性較低（資訊透明度矯正管理者過度樂觀；"
          "預期 β3 顯著為正）。", indent=False)

h(doc, "五、初步結果（AI 輔助驗證）", 2)
para(doc, "1. 樣本存在穩健的成本黏性：核心係數 β2（D×ΔLNREV）在水回收率樣本顯著為負"
          "（全樣本 −1.27***、高耗水產業 −1.48***、遞延兩期 −1.15***），"
          "顯示營收下降時費用向下調整較不敏感，與文獻一致。")
para(doc, "2. 水管理的調節效果（β3）未獲支持：無論當期、遞延（t-1／t-2），或在"
          "高／低耗水產業子樣本中，H1 與 H2 的核心三重交乘 β3 均「未達統計顯著」。")
para(doc, "3. 控制變數方向符合預期（資產密集度、ROA、連續兩年下降等之交乘項多顯著），"
          "顯示模型設定與資料品質可信。")

h(doc, "六、洞察", 2)
bullet(doc, "台灣製造業存在穩健成本黏性，資料與模型可信，可作為研究基礎。")
bullet(doc, "β3 不顯著的可能原因偏向「衡量與統計功效」而非必然無關係：水回收率有效"
            "樣本僅約千筆且變異有限，水揭露多為二元、資訊量低。")
bullet(doc, "值得注意：以較粗產業分類時曾出現邊際顯著，改用精準 TSE 產業定義後即消失，"
            "顯示該效果並不穩健，後續須審慎，避免過度解讀。")

h(doc, "七、後續可行方向（若老師認為方向可行）", 2)
bullet(doc, "強化水衡量：改以 GRI 用水揭露度（連續）、用水密集度（總用水量／營收）、"
            "回收利用水量／總用水量等建構水管理指數，擴大樣本與變異。")
bullet(doc, "識別策略：以 2021 年台灣乾旱（限水）作為外生水資源衝擊，採差異中的差異"
            "（DiD）檢驗水管理較佳者是否較能緩衝營運衝擊。")
bullet(doc, "或調整應變數：改看水管理對營業費用率、獲利或營運風險之直接影響，"
            "以提升檢定力。")

h(doc, "八、請示", 2)
para(doc, "懇請老師指示本主題與研究設計是否可行。若老師認為方向可行，我將以現有可"
          "重複執行之資料管線為基礎，優先補強水衡量與識別策略後持續深化；"
          "亦歡迎老師就變數或模型設定提供進一步建議。")

# ---- 附頁：結果摘要表（18 個設定）----
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
h(doc, "附表：全部設定之 β2／β3 對照（18 個設定）", 2)
para(doc, "下表彙整「水衡量 × 產業組 × 遞延期」共 18 個設定之核心係數。β2 為整體成本"
          "黏性（D×ΔLNREV），預期為負；β3 為水管理之三重交乘（Water×D×ΔLNREV），"
          "即 H1／H2 之核心係數。星號 *、**、*** 分別代表 10%、5%、1% 顯著水準；"
          "無星號者不顯著。估計皆採 OLS＋產業與年份固定效果＋公司叢集穩健標準誤。",
     indent=False)
add_summary_table(doc, "06_robustness_summary.csv")
para(doc, "說明：β2 於水回收率相關設定（尤其高耗水產業當期）穩健顯著為負，確認成本"
          "黏性存在；β3 於所有 18 個設定均不顯著，故現階段 H1／H2 未獲支持。",
     indent=False)

doc.save(OUT)
print("已生成：", OUT)
