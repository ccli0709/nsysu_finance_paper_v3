# -*- coding: utf-8 -*-
"""
01_merge_tej_data.py
--------------------
把 tej_data/ 內的 TEJ zip（單一 CSV）合併成「證券代碼 + 西元年份」的年度面板。

三個原始資料集（更新資料時只需改下方 *_ZIP 檔名與欄位清單）：
  1. 財務（單季，年月 MM=03/06/09/12）→ 年度：流量加總、存量取年末 Q4。
  2. 水資源（年度）→ 直接年度化。
  3. ESG（年月 MM=06/12）→ 年度化（優先取 12）；SASB主產業視為公司靜態屬性。

輸出：01_merged_tej_data.csv（UTF-8-SIG）。
"""

import os
import io
import zipfile
import pandas as pd
import numpy as np

# ----------------------------------------------------------------------------
# 設定：原始資料集檔名與欄位（資料更新時在此登記）
# ----------------------------------------------------------------------------
DATA_DIR = "tej_data"
OUTPUT_CSV = "01_merged_tej_data.csv"

FIN_ZIP = "TEJ20260723012741.zip"      # IFRS 以合併為主財務(單季)-一般產業Ⅳ
WATER_ZIP = "TEJ20260723014020.zip"    # TESG 環境構面-企業用水量（年度）
ESG_ZIP = "TEJ20260723104157.zip"      # TESG ESG-變數彙總表（年月）
TSE_ZIP = "TEJ20260724105926.zip"      # TEJ Company DB-屬性基本資料（TSE/TEJ 產業別，靜態）
WASTE_ESG_ZIP = "TEJ20260727061428.zip"   # TESG ESG-廢棄物揭露（年月）
WASTE_QTY_ZIP = "TEJ20260727062055.zip"   # TESG-廢棄物量/密集度/認證（年度）
WASTE_FINE_ZIP = "TEJ20260727063019.zip"  # TESG-廢棄物罰鍰次數（年度）

# 財務：流量變數（同年四季加總，需四季齊全）
FIN_FLOW_COLS = ["營業費用", "推銷費用", "管理費用", "研究發展費", "營業收入淨額"]
# 財務：存量／比率變數（取年末 Q4 該季值）
FIN_STOCK_COLS = [
    "資產總額", "員工人數-母公司",
    "ROA(C)稅前息前折舊前", "ROA(A)稅後息前",
    "負債比率", "總負債/總淨值",
]
REQUIRE_FULL_YEAR = True   # 流量需四季齊全才視為有效年度值

# 水資源：年度值欄位（同年多筆序號取首個非空值）
WATER_VALUE_COLS = ["回收利用水量", "製程水回收率%", "水回收率%", "總用水量(含回收水)"]

# ESG：年度揭露欄（以年份 left join）
ESG_ANNUAL_COLS = ["E_用水回收率%", "E_取得用水相關國際認證", "E_GRI_用水及廢水管理揭露度"]
# ESG：公司靜態屬性
ESG_STATIC_COLS = ["SASB主產業"]

# TSE 產業別（公司靜態，以 [證券代碼] 套用）
TSE_STATIC_COLS = ["TSE產業_代碼", "TSE產業_名稱", "TEJ產業_名稱", "TEJ子產業_名稱"]

# 廢棄物 ESG 揭露（年月，優先取年末 12）
WASTE_ESG_COLS = ["E_GRI_廢棄物管理揭露度", "E_每百萬營收廢棄物",
                  "E_有害廢棄物回收百分比%", "G_報廢產品及電子廢棄物再循環百分比%"]
# 廢棄物量/密集度（年度，含自由文字欄，需 robust 讀取；僅取數值欄）
WASTE_QTY_COLS = ["總重量(有害+非有害)(噸)", "廢棄物密集度(公噸/單位)",
                  "有害廢棄物量(公噸)", "非有害廢棄物量(公噸)"]
# 廢棄物罰鍰次數（年度）
WASTE_FINE_COLS = ["一般廢棄物罰鍰次數", "事業廢棄物罰鍰次數", "廢棄物清除處理機構罰鍰次數"]


# ----------------------------------------------------------------------------
# 工具函式
# ----------------------------------------------------------------------------
def read_tej_zip(zip_name, robust=False):
    """讀取 TEJ zip 內單一 CSV（UTF-16 LE、Tab 分隔），回傳 DataFrame。
    robust=True 時以 python 引擎並跳過欄數異常列（適用含自由文字、內嵌 Tab 的檔案）。"""
    path = os.path.join(DATA_DIR, zip_name)
    with zipfile.ZipFile(path) as z:
        csv_name = [n for n in z.namelist() if n.lower().endswith(".csv")][0]
        raw = z.read(csv_name)
    if robust:
        df = pd.read_csv(io.BytesIO(raw), encoding="utf-16", sep="\t", dtype=str,
                         engine="python", on_bad_lines="skip")
    else:
        df = pd.read_csv(io.BytesIO(raw), encoding="utf-16", sep="\t", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    return df


def clean_code(series):
    """證券代碼 '2201 裕隆' → '2201'（取第一個空白前的代碼，保留前導 0）。"""
    return series.astype(str).str.strip().str.split(n=1).str[0]


def to_num(series):
    """去除逗號/空白後轉數值。"""
    s = series.astype(str).str.replace(",", "", regex=False).str.strip()
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    return pd.to_numeric(s, errors="coerce")


def yyyymm_to_year_month(series):
    """年月字串（如 200506）→ (year, month) 整數 Series；非法值為 NaN。"""
    ym = pd.to_numeric(series, errors="coerce")
    year = np.floor(ym / 100).astype("Int64")
    month = (ym % 100).astype("Int64")
    return year, month


# ----------------------------------------------------------------------------
# 1) 財務（單季）→ 年度
# ----------------------------------------------------------------------------
def build_financial_annual():
    fin = read_tej_zip(FIN_ZIP)
    fin["證券代碼"] = clean_code(fin["證券代碼"])
    year, month = yyyymm_to_year_month(fin["年月"])
    fin["西元年份"] = year
    fin["月"] = month
    fin = fin[fin["西元年份"].notna() & fin["月"].notna()].copy()
    fin["西元年份"] = fin["西元年份"].astype(int)
    fin["月"] = fin["月"].astype(int)

    for c in FIN_FLOW_COLS + FIN_STOCK_COLS:
        if c in fin.columns:
            fin[c] = to_num(fin[c])

    grp = fin.groupby(["證券代碼", "西元年份"])

    # 流量：四季加總（需每欄四季齊全）
    flow_sum = grp[FIN_FLOW_COLS].sum(min_count=1)
    if REQUIRE_FULL_YEAR:
        flow_count = grp[FIN_FLOW_COLS].count()
        flow_annual = flow_sum.where(flow_count == 4)
    else:
        flow_annual = flow_sum

    # 存量／比率：取年末 Q4（12 月）
    q4 = fin[fin["月"] == 12]
    stock_annual = q4.groupby(["證券代碼", "西元年份"])[FIN_STOCK_COLS].first()

    fin_annual = flow_annual.join(stock_annual, how="outer").reset_index()
    print(f"[財務] 原始 {len(fin):,} 季列 → 年度 {len(fin_annual):,} 列")
    return fin_annual


# ----------------------------------------------------------------------------
# 2) 水資源（年度）
# ----------------------------------------------------------------------------
def build_water_annual():
    water = read_tej_zip(WATER_ZIP)
    water["證券代碼"] = clean_code(water["證券代碼"])
    water["西元年份"] = pd.to_numeric(water["年"], errors="coerce").astype("Int64")
    water = water[water["西元年份"].notna()].copy()
    water["西元年份"] = water["西元年份"].astype(int)

    for c in WATER_VALUE_COLS:
        if c in water.columns:
            water[c] = to_num(water[c])

    # 同年多筆序號 → 取每欄首個非空值
    water_annual = (
        water.groupby(["證券代碼", "西元年份"])[WATER_VALUE_COLS].first().reset_index()
    )
    print(f"[水資源] 原始 {len(water):,} 列 → 年度 {len(water_annual):,} 列")
    return water_annual


# ----------------------------------------------------------------------------
# 3) ESG（年月 06/12）→ 年度 + 公司靜態產業別
# ----------------------------------------------------------------------------
def build_esg_annual():
    esg = read_tej_zip(ESG_ZIP)
    esg["證券代碼"] = clean_code(esg["證券代碼"])
    year, month = yyyymm_to_year_month(esg["年月"])
    esg["西元年份"] = year
    esg["月"] = month
    esg = esg[esg["西元年份"].notna() & esg["月"].notna()].copy()
    esg["西元年份"] = esg["西元年份"].astype(int)
    esg["月"] = esg["月"].astype(int)

    for c in ESG_ANNUAL_COLS:
        if c in esg.columns:
            esg[c] = to_num(esg[c])

    # 每 (代碼, 年) 取月份最大者（優先年末 12），保留整列
    idx = esg.groupby(["證券代碼", "西元年份"])["月"].idxmax()
    esg_year = esg.loc[idx]
    esg_annual = esg_year[["證券代碼", "西元年份"] + ESG_ANNUAL_COLS].reset_index(drop=True)

    # SASB主產業：公司靜態，取最新年度的非空值
    static_src = esg.dropna(subset=ESG_STATIC_COLS).sort_values(
        ["證券代碼", "西元年份", "月"]
    )
    esg_static = static_src.groupby("證券代碼")[ESG_STATIC_COLS].last()

    print(f"[ESG] 原始 {len(esg):,} 列 → 年度 {len(esg_annual):,} 列；"
          f"有產業別公司 {len(esg_static):,} 家")
    return esg_annual, esg_static


# ----------------------------------------------------------------------------
# 4) TEJ Company DB 產業別（公司靜態）
# ----------------------------------------------------------------------------
def build_tse_static():
    tse = read_tej_zip(TSE_ZIP)
    tse["證券代碼"] = clean_code(tse["證券代碼"])
    cols = [c for c in TSE_STATIC_COLS if c in tse.columns]
    for c in cols:
        tse[c] = tse[c].astype(str).str.strip().replace(
            {"": np.nan, "nan": np.nan, "None": np.nan})
    src = tse.dropna(subset=cols, how="all")
    tse_static = src.groupby("證券代碼")[cols].first()
    print(f"[TSE產業] 原始 {len(tse):,} 列 → 靜態產業別公司 {len(tse_static):,} 家")
    return tse_static


# ----------------------------------------------------------------------------
# 5) 廢棄物管理指標（三個資料集）
# ----------------------------------------------------------------------------
def build_waste_esg_annual():
    """廢棄物 ESG 揭露（年月 06/12 → 年度，優先取 12）。"""
    esg = read_tej_zip(WASTE_ESG_ZIP)
    esg["證券代碼"] = clean_code(esg["證券代碼"])
    year, month = yyyymm_to_year_month(esg["年月"])
    esg["西元年份"], esg["月"] = year, month
    esg = esg[esg["西元年份"].notna() & esg["月"].notna()].copy()
    esg["西元年份"] = esg["西元年份"].astype(int)
    esg["月"] = esg["月"].astype(int)
    cols = [c for c in WASTE_ESG_COLS if c in esg.columns]
    for c in cols:
        esg[c] = to_num(esg[c])
    idx = esg.groupby(["證券代碼", "西元年份"])["月"].idxmax()
    out = esg.loc[idx, ["證券代碼", "西元年份"] + cols].reset_index(drop=True)
    print(f"[廢棄物揭露] 原始 {len(esg):,} 列 → 年度 {len(out):,} 列")
    return out


def build_waste_qty_annual():
    """廢棄物量/密集度（年度，robust 讀取，僅取數值欄）。"""
    q = read_tej_zip(WASTE_QTY_ZIP, robust=True)
    q["證券代碼"] = clean_code(q["證券代碼"])
    q["西元年份"] = pd.to_numeric(q["年"], errors="coerce").astype("Int64")
    q = q[q["西元年份"].notna()].copy()
    q["西元年份"] = q["西元年份"].astype(int)
    cols = [c for c in WASTE_QTY_COLS if c in q.columns]
    for c in cols:
        q[c] = to_num(q[c])
    out = q.groupby(["證券代碼", "西元年份"])[cols].first().reset_index()
    print(f"[廢棄物量] 原始 {len(q):,} 列 → 年度 {len(out):,} 列")
    return out


def build_waste_fine_annual():
    """廢棄物罰鍰次數（年度）。"""
    f = read_tej_zip(WASTE_FINE_ZIP)
    f["證券代碼"] = clean_code(f["證券代碼"])
    f["西元年份"] = pd.to_numeric(f["年"], errors="coerce").astype("Int64")
    f = f[f["西元年份"].notna()].copy()
    f["西元年份"] = f["西元年份"].astype(int)
    cols = [c for c in WASTE_FINE_COLS if c in f.columns]
    for c in cols:
        f[c] = to_num(f[c])
    out = f.groupby(["證券代碼", "西元年份"])[cols].max().reset_index()
    print(f"[廢棄物罰鍰] 原始 {len(f):,} 列 → 年度 {len(out):,} 列")
    return out


# ----------------------------------------------------------------------------
# 合併
# ----------------------------------------------------------------------------
def main():
    fin_annual = build_financial_annual()
    water_annual = build_water_annual()
    esg_annual, esg_static = build_esg_annual()
    tse_static = build_tse_static()
    waste_esg = build_waste_esg_annual()
    waste_qty = build_waste_qty_annual()
    waste_fine = build_waste_fine_annual()

    # 財務 ⟗ 水資源（outer join，保留兩邊）
    merged = fin_annual.merge(water_annual, on=["證券代碼", "西元年份"], how="outer")
    # ⟕ ESG 年度揭露欄
    merged = merged.merge(esg_annual, on=["證券代碼", "西元年份"], how="left")
    # ⟕ 廢棄物三個資料集（年度 left join）
    merged = merged.merge(waste_esg, on=["證券代碼", "西元年份"], how="left")
    merged = merged.merge(waste_qty, on=["證券代碼", "西元年份"], how="left")
    merged = merged.merge(waste_fine, on=["證券代碼", "西元年份"], how="left")
    # 套用公司靜態 SASB主產業
    for c in ESG_STATIC_COLS:
        merged[c] = merged["證券代碼"].map(esg_static[c])
    # 套用公司靜態 TSE/TEJ 產業別
    for c in tse_static.columns:
        merged[c] = merged["證券代碼"].map(tse_static[c])

    merged = merged.sort_values(["證券代碼", "西元年份"]).reset_index(drop=True)
    merged.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    # 摘要
    print("\n===== 合併結果摘要 =====")
    print(f"輸出檔：{OUTPUT_CSV}")
    print(f"總列數：{len(merged):,}")
    print(f"公司數：{merged['證券代碼'].nunique():,}")
    print(f"年份範圍：{int(merged['西元年份'].min())} - {int(merged['西元年份'].max())}")
    print("關鍵欄非空數：")
    for c in ["營業費用", "營業收入淨額", "資產總額", "水回收率%",
              "製程水回收率%", "回收利用水量", "E_GRI_用水及廢水管理揭露度",
              "SASB主產業", "TSE產業_名稱",
              "E_每百萬營收廢棄物", "E_GRI_廢棄物管理揭露度", "事業廢棄物罰鍰次數",
              "廢棄物密集度(公噸/單位)"]:
        if c in merged.columns:
            print(f"  {c:<22}: {merged[c].notna().sum():,}")


if __name__ == "__main__":
    main()
