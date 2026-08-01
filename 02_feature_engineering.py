# -*- coding: utf-8 -*-
"""
02_feature_engineering.py
-------------------------
從 01_merged_tej_data.csv 建構成本黏性 (cost stickiness) 研究所需變數。

依 ABJ (Anderson, Banker & Janakiraman 2003) 與 study.md 的設計：
  應變數/核心：ΔLNSGA、ΔLNREV、D（收入下降虛擬）
  自變數（水管理）：Water_Rate（連續）、Water_Disc（揭露虛擬）
  控制變數：Size、AI、EI、ROA、Lev、Decrease
  對連續變數做前後 1% 縮尾（_w 欄）。

年度差分變數（ΔLN、D）僅在「連續年份 (t 與 t-1 相鄰)」時計算，否則為缺值。

輸出：02_features_data.csv（UTF-8-SIG）。
"""

import numpy as np
import pandas as pd

INPUT_CSV = "01_merged_tej_data.csv"
OUTPUT_CSV = "02_features_data.csv"

# 營業費用衡量：主分析用「營業費用」欄；穩健性另建「推銷費用+管理費用」
SGA_MAIN_COL = "營業費用"
REV_COL = "營業收入淨額"

# ROA 主用稅後息前 ROA(A)；ROA(C) 保留作穩健性
ROA_MAIN_COL = "ROA(A)稅後息前"

# 縮尾比例（前後各 1%）
WINSOR_P = 0.01
# 需縮尾的連續變數
WINSOR_COLS = [
    "dLNSGA", "dLNSGA_alt", "dLNREV",
    "Water_Rate", "水回收率%", "製程水回收率%",
    "Waste_Intensity", "Waste_Density", "Waste_Fine",
    "Waste_FineInt", "Water_Intensity",
    "Size", "AI", "EI", "ROA", "Lev",
]


def safe_ln(s):
    """僅對正值取自然對數，其餘為 NaN。"""
    s = pd.to_numeric(s, errors="coerce")
    return np.log(s.where(s > 0))


def winsorize(s, p=WINSOR_P):
    s = pd.to_numeric(s, errors="coerce")
    lo, hi = s.quantile(p), s.quantile(1 - p)
    return s.clip(lower=lo, upper=hi)


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig")
    df["證券代碼"] = df["證券代碼"].astype(str)
    df = df.sort_values(["證券代碼", "西元年份"]).reset_index(drop=True)

    g = df.groupby("證券代碼")
    prev_year = g["西元年份"].shift(1)
    consec = (df["西元年份"] - prev_year) == 1   # t 與 t-1 是否相鄰年

    def lag(col):
        """取同公司前一年度值，僅在相鄰年份有效。"""
        return g[col].shift(1).where(consec)

    # 推銷費用 + 管理費用（替代 SG&A 衡量）
    df["SGA_alt"] = pd.to_numeric(df["推銷費用"], errors="coerce").fillna(0) + \
                    pd.to_numeric(df["管理費用"], errors="coerce").fillna(0)
    df.loc[df["推銷費用"].isna() & df["管理費用"].isna(), "SGA_alt"] = np.nan

    # ---- 應變數與核心變數 ----
    sga = pd.to_numeric(df[SGA_MAIN_COL], errors="coerce")
    rev = pd.to_numeric(df[REV_COL], errors="coerce")
    df["dLNSGA"] = safe_ln(sga) - safe_ln(lag(SGA_MAIN_COL))
    df["dLNSGA_alt"] = safe_ln(df["SGA_alt"]) - safe_ln(lag("SGA_alt"))
    df["dLNREV"] = safe_ln(rev) - safe_ln(lag(REV_COL))

    # 收入下降虛擬變數 D（本期營收 < 前期）
    rev_lag = lag(REV_COL)
    both = rev.notna() & rev_lag.notna()
    df["D"] = np.where(both, (rev < rev_lag).astype(float), np.nan)

    # 連續兩年營收下降 Decrease（D_t=1 且 D_{t-1}=1）
    d_lag = lag("D")
    df["Decrease"] = np.where(
        df["D"].notna() & d_lag.notna(),
        ((df["D"] == 1) & (d_lag == 1)).astype(float),
        np.nan,
    )

    # ---- 自變數：水管理 ----
    df["Water_Rate"] = pd.to_numeric(df["水回收率%"], errors="coerce")  # 主分析衡量
    # 做法 A：該公司當年有水量/水回收紀錄則揭露=1
    has_water = (
        pd.to_numeric(df["回收利用水量"], errors="coerce").notna()
        | pd.to_numeric(df["水回收率%"], errors="coerce").notna()
        | pd.to_numeric(df["製程水回收率%"], errors="coerce").notna()
    )
    df["Water_Disc"] = has_water.astype(float)          # 做法 A（主）
    # 做法 B：GRI 用水及廢水管理揭露度 > 0
    gri = pd.to_numeric(df["E_GRI_用水及廢水管理揭露度"], errors="coerce")
    df["Water_Disc_GRI"] = (gri.fillna(0) > 0).astype(float)  # 做法 B（穩健性）

    # ---- 自變數：廢棄物管理（延伸主題）----
    # 主分析（實質投入）：每百萬營收廢棄物量 → 廢棄物處理合約/費用之僵固性
    df["Waste_Intensity"] = pd.to_numeric(df["E_每百萬營收廢棄物"], errors="coerce")
    # 次分析（資訊透明度）：GRI 廢棄物管理揭露度（連續，揭露品質）
    df["Waste_Disc"] = pd.to_numeric(df["E_GRI_廢棄物管理揭露度"], errors="coerce")
    # 穩健/風險衝擊：事業廢棄物罰鍰次數（環境違規事件）
    df["Waste_Fine"] = pd.to_numeric(df["事業廢棄物罰鍰次數"], errors="coerce")
    # 補充：廢棄物密集度（公噸/單位，來源資料集2，覆蓋較小）
    df["Waste_Density"] = pd.to_numeric(df["廢棄物密集度(公噸/單位)"], errors="coerce")

    # 新增（v5）：廢棄物裁罰金額（連續，取代罰鍰次數以解決極端值與離散問題）
    assets_ = pd.to_numeric(df["資產總額"], errors="coerce")  # 千元
    fine_amt = pd.to_numeric(df["廢棄物裁罰金額"], errors="coerce")  # 元
    in_fine_db = pd.to_numeric(df["事業廢棄物罰鍰次數"], errors="coerce").notna()
    # 在裁罰資料庫涵蓋範圍內、但無廢棄物裁罰者視為 0
    fine_amt = fine_amt.where(fine_amt.notna(), other=np.where(in_fine_db, 0.0, np.nan))
    df["Waste_FineAmt"] = fine_amt
    # 佔資產比（每百萬元資產之廢棄物裁罰金額；資產千元→元 ×1000）
    df["Waste_FineInt"] = fine_amt / (assets_ * 1000.0) * 1e6
    # 用水密集度（總用水量含回收水 / 營收），擴大水構面樣本
    total_water = pd.to_numeric(df["總用水量(含回收水)"], errors="coerce")
    df["Water_Intensity"] = total_water / rev.where(rev > 0)

    # ---- 控制變數 ----
    assets = pd.to_numeric(df["資產總額"], errors="coerce")
    emp = pd.to_numeric(df["員工人數-母公司"], errors="coerce")
    df["Size"] = safe_ln(assets)                       # LN(資產總額)
    df["AI"] = assets / rev.where(rev > 0)             # 資產密集度
    df["EI"] = emp / rev.where(rev > 0)                # 員工密集度
    df["ROA"] = pd.to_numeric(df[ROA_MAIN_COL], errors="coerce")
    df["Lev"] = pd.to_numeric(df["負債比率"], errors="coerce")

    # ---- 縮尾 Winsorize（_w 欄）----
    for c in WINSOR_COLS:
        if c in df.columns:
            df[c + "_w"] = winsorize(df[c])

    # ---- 遞延期水管理變數（方向四：時間落差效應，不新增概念變數，僅平移）----
    # 將既有水變數平移 t-1、t-2；僅在年份差恰為 k 時有效（避免跨越缺漏年份）。
    gl = df.groupby("證券代碼")

    def lag_k(col, k):
        shifted = gl[col].shift(k)
        yr_shift = gl["西元年份"].shift(k)
        valid = (df["西元年份"] - yr_shift) == k
        return shifted.where(valid)

    LAG_TARGETS = ["Water_Rate_w", "製程水回收率%_w", "Water_Disc", "Water_Disc_GRI",
                   "Waste_Intensity_w", "Waste_Disc", "Waste_Fine_w",
                   "Waste_FineInt_w", "Water_Intensity_w"]
    for col in LAG_TARGETS:
        for k in (1, 2):
            df[f"{col}_l{k}"] = lag_k(col, k)

    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    # 摘要
    print("===== 變數建構摘要 =====")
    print(f"輸出檔：{OUTPUT_CSV}")
    print(f"總列數：{len(df):,}；公司數：{df['證券代碼'].nunique():,}")
    key = ["dLNSGA", "dLNREV", "D", "Decrease", "Water_Rate", "Water_Disc",
           "Waste_Intensity", "Waste_Disc", "Waste_Fine",
           "Size", "AI", "EI", "ROA", "Lev"]
    print("關鍵變數非空數：")
    for c in key:
        print(f"  {c:<16}: {df[c].notna().sum():,}")
    # 迴歸最低要求（三者皆有）
    core_mask = df[["dLNSGA", "dLNREV", "D"]].notna().all(axis=1)
    print(f"\n三項核心變數(dLNSGA/dLNREV/D)皆非空：{core_mask.sum():,} 列")
    print(f"其中 Water_Rate 非空：{df.loc[core_mask, 'Water_Rate'].notna().sum():,} 列")
    print("遞延期水變數非空數（核心樣本內）：")
    for c in ["Water_Rate_w_l1", "Water_Rate_w_l2",
              "Water_Disc_l1", "Water_Disc_l2"]:
        print(f"  {c:<18}: {df.loc[core_mask, c].notna().sum():,}")


if __name__ == "__main__":
    main()
