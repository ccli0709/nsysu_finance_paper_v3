# -*- coding: utf-8 -*-
"""從「台灣碩博士論文題目清單」挑選 30 筆相關碩博論文，輸出 Word（含連結與 H1–H5/機制對應）。"""
import pandas as pd, glob, os, re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

FOLDER = "台灣碩博士論文題目清單"
OUT = "參考文獻候選清單_水與廢棄物成本黏性.docx"
CJK, EN = "新細明體", "Times New Roman"

# ---------------- 讀取 ----------------
rows = []
for f in glob.glob(os.path.join(FOLDER, "fb260621_ESG-114-*.csv")):
    d = pd.read_csv(f, encoding="utf-8-sig", dtype=str)
    for _, r in d.iterrows():
        rows.append({"year": 114, "title": str(r.get("論文名稱", "")).strip(),
                     "author": str(r.get("研究生", "")).strip(),
                     "school": str(r.get("校院名稱", "")).strip(),
                     "dept": str(r.get("系所名稱", "")).strip(),
                     "url": str(r.get("引用網址", "")).strip()})
for f in glob.glob(os.path.join(FOLDER, "*ndltd.csv")):
    yr = int(re.match(r"(\d+)ndltd", os.path.basename(f)).group(1))
    d = pd.read_csv(f, encoding="utf-8-sig", dtype=str)
    for _, r in d.iterrows():
        rows.append({"year": yr, "title": str(r.get("論文名稱(中文)", "")).strip(),
                     "author": str(r.get("作者", "")).strip(),
                     "school": "", "dept": "",
                     "url": str(r.get("博碩士論文網址", "")).strip()})
df = pd.DataFrame(rows)
df = df[df["title"].str.len() > 0].drop_duplicates(subset=["title", "author"])

# ---------------- 分類與排除 ----------------
CHEM = ["合成", "氫化", "感測", "微量天平", "黏度", "呋喃", "螺環", "晶體", "反應於"]

def is_chem(t):
    return any(k in t for k in CHEM)

def is_cost_stick(t):
    return bool(re.search("成本僵固|成本黏|費用黏|成本不對稱|成本不對稱", t)) and not is_chem(t)

def is_env_cost(t):
    env = any(k in t for k in ["環境", "ESG", "永續", "碳排", "碳揭露", "水", "廢棄物",
                                "污染", "綠色", "揭露", "循環經濟", "氣候", "永續報告"])
    # 去除僅談「績效」之通用題；要求較實質之成本/風險/揭露/估值等財務連結
    fin_strong = any(k in t for k in ["成本", "資金成本", "代理成本", "風險", "投資決策",
                                       "融資", "估值", "價值攸關", "負債", "權益",
                                       "盈餘管理", "漂綠", "資源密集", "債務", "資本"])
    # 明顯偏投資/基金/股價策略者排除
    noise = any(k in t for k in ["ETF", "指數型基金", "基金", "羊群", "本益比",
                                  "投資策略", "股票投資", "股價異常"])
    return env and fin_strong and not noise and not is_chem(t)

def mechanism(t):
    if is_cost_stick(t):
        if any(k in t for k in ["ESG", "環境", "永續", "碳", "廢棄物", "水", "綠色"]):
            return "環境×成本黏性（H1/H3 實質投入；核心對照）"
        return "成本黏性方法論／β2 基準（依變數與模型）"
    if any(k in t for k in ["廢棄物", "水", "污染"]):
        return "H3/H5 廢棄物·實質投入與違規風險"
    if any(k in t for k in ["揭露", "資訊", "透明", "漂綠"]):
        return "H2/H4 資訊揭露機制（透明度／監督）"
    if any(k in t for k in ["碳排", "碳揭露", "氣候", "永續報告"]):
        return "環境績效／揭露→成本（H4 機制、高污染情境）"
    if "代理成本" in t or "風險" in t:
        return "代理成本／風險（β3 調節機制）"
    return "ESG→成本／績效（背景與假設支撐）"

df["chem"] = df["title"].apply(is_chem)
df = df[~df["chem"]]
df["cost_stick"] = df["title"].apply(is_cost_stick)
df["env_cost"] = df["title"].apply(is_env_cost)

# ---------------- 挑選 30 ----------------
# 「金牌」最貼題（水/廢棄物 或 環境揭露→資金成本 或 環境×成本黏性），不論年份必納入
def is_star(t):
    if is_cost_stick(t) and any(k in t for k in ["環境", "ESG", "永續", "碳", "廢棄物", "水", "綠色"]):
        return True
    if any(k in t for k in ["廢棄物", "水"]) and any(k in t for k in ["法規", "投資決策", "融資", "成本"]):
        return True
    if "氣候環境資訊揭露" in t:
        return True
    if any(k in t for k in ["碳排", "高碳排", "永續報告"]) and any(k in t for k in ["資金成本", "負債", "成本"]):
        return True
    return False

df["star"] = df["title"].apply(is_star)
df["cost_stick"] = df["title"].apply(is_cost_stick)
df["priority"] = (df["star"] | df["cost_stick"]).astype(int)

core = df[df["cost_stick"]]
star = df[df["star"] & ~df["cost_stick"]]
env = df[df["env_cost"] & ~df["priority"].astype(bool)]

picked = pd.concat([core, star, env.sort_values("year", ascending=False).head(30)])
picked = picked.drop_duplicates(subset=["title", "author"])
picked["is_core"] = picked["cost_stick"].astype(int)
# 先確保高優先（金牌+成本黏性），再依年份新→舊；取 30
picked = picked.sort_values(["priority", "year", "is_core"], ascending=[False, False, False]).head(30)
picked = picked.sort_values(["priority", "year", "is_core"], ascending=[False, False, False]).reset_index(drop=True)

print("挑選筆數:", len(picked), "各年:", dict(picked["year"].value_counts().sort_index()))

# ---------------- 產生 Word ----------------
def cjk(run):
    run.font.name = EN
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

doc = Document()
n = doc.styles["Normal"]; n.font.name = EN; n.font.size = Pt(11)
n.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("參考文獻候選清單：水與廢棄物管理對成本黏性之影響"); r.bold = True
r.font.size = Pt(15); cjk(r)
s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.CENTER
cjk(s.add_run("資料來源：臺灣博碩士論文知識加值系統（NDLTD）題目清單；依學年度新到舊、成本黏性核心優先"))

p = doc.add_paragraph()
cjk(p.add_run(f"共 {len(picked)} 筆。「對應」欄標示與本研究假設（H1 水績效、H2 水揭露、"
              "H3 廢棄物密集度、H4 廢棄物揭露、H5 廢棄物罰鍰）或文獻機制之關聯。"))

cols = ["#", "學年度", "論文題目", "作者", "校院/系所", "對應 H1–H5／機制", "連結"]
tb = doc.add_table(rows=1, cols=len(cols)); tb.style = "Light Grid Accent 1"
for j, c in enumerate(cols):
    rr = tb.rows[0].cells[j].paragraphs[0].add_run(c); rr.bold = True
    rr.font.size = Pt(9); cjk(rr)
for i, row in picked.iterrows():
    school = (row["school"] + (" " + row["dept"] if row["dept"] else "")).strip() or "—"
    vals = [str(i + 1), f"{row['year']}", row["title"], row["author"] or "—",
            school, mechanism(row["title"]), row["url"] or "—"]
    cells = tb.add_row().cells
    for j, v in enumerate(vals):
        rr = cells[j].paragraphs[0].add_run(str(v)); rr.font.size = Pt(8); cjk(rr)

doc.save(OUT)
print("已生成：", OUT)
for i, row in picked.iterrows():
    print(f"[{row['year']}] {row['title']} /{row['author']} -> {mechanism(row['title'])}")
