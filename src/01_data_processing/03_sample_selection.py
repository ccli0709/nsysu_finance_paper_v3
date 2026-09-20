# -*- coding: utf-8 -*-
"""
03_sample_selection.py
----------------------
從 02_features_data.csv 篩出成本黏性迴歸樣本，並輸出各產業資料量統計。

篩選（參數見下）：
  1. 限定研究期間 START_YEAR ~ END_YEAR。
  2. 排除金融保險業（SASB主產業 含「金融」）。
  3. 剔除缺產業別者（產業固定效果需要）。
  4. 剔除核心變數 dLNSGA / dLNREV / D 任一缺值。
     （營收或營業費用 ≤ 0 者在 02 取對數時已成 NaN，於此一併剔除。）

輸出：03_sample_data.csv、03_industry_stats.csv（皆 UTF-8-SIG）。
"""

import numpy as np
import pandas as pd

INPUT_CSV = "data/processed/02_features_data.csv"
OUTPUT_SAMPLE = "data/processed/03_sample_data.csv"
OUTPUT_STATS = "data/processed/03_industry_stats.csv"

# 篩選參數
START_YEAR = 2014
END_YEAR = 2024
EXCLUDE_INDUSTRY_KEYWORD = "金融"     # SASB主產業 含此字串者排除
DROP_MISSING_INDUSTRY = True          # 剔除無產業別者
CORE_COLS = ["dLNSGA", "dLNREV", "DEC"]
CONTROL_COLS = ["SIZE", "AI", "EI", "ROA", "LEV", "SUCC_DEC"]


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)
    n0 = len(df)

    # 1) 研究期間
    df = df[(df["西元年份"] >= START_YEAR) & (df["西元年份"] <= END_YEAR)]
    n1 = len(df)

    # 2) 排除金融保險業
    ind = df["SASB主產業"].astype(str)
    df = df[~ind.str.contains(EXCLUDE_INDUSTRY_KEYWORD, na=False)]
    n2 = len(df)

    # 3) 剔除缺產業別
    if DROP_MISSING_INDUSTRY:
        df = df[df["SASB主產業"].notna() & (df["SASB主產業"].astype(str).str.strip() != "")]
    n3 = len(df)

    # 4) 剔除核心變數缺值
    df = df[df[CORE_COLS].notna().all(axis=1)]
    n4 = len(df)

    df = df.sort_values(["證券代碼", "西元年份"]).reset_index(drop=True)
    df.to_csv(OUTPUT_SAMPLE, index=False, encoding="utf-8-sig")

    # ---- 產業統計 ----
    reg_ok = df[CORE_COLS + CONTROL_COLS].notna().all(axis=1)
    tmp = df.copy()
    tmp["_reg_ok"] = reg_ok.astype(int)
    tmp["_has_water"] = tmp["WATER_RATE"].notna().astype(int)
    tmp["_disc"] = (tmp["WATER_DISC"] == 1).astype(int)
    tmp["_reg_water"] = (reg_ok & tmp["WATER_RATE"].notna()).astype(int)

    stats = tmp.groupby("SASB主產業").agg(
        公司數=("證券代碼", "nunique"),
        觀測數=("證券代碼", "size"),
        有水資料筆數=("_has_water", "sum"),
        有揭露筆數=("_disc", "sum"),
        迴歸可用筆數=("_reg_ok", "sum"),
        迴歸且有水筆數=("_reg_water", "sum"),
    ).sort_values("觀測數", ascending=False).reset_index()
    stats.to_csv(OUTPUT_STATS, index=False, encoding="utf-8-sig")

    # 摘要
    print("===== 樣本篩選摘要 =====")
    print(f"原始              : {n0:,}")
    print(f"限研究期間 {START_YEAR}-{END_YEAR} : {n1:,}")
    print(f"排除金融業後      : {n2:,}")
    print(f"剔除缺產業別後    : {n3:,}")
    print(f"核心變數皆非空後  : {n4:,}  ← 輸出樣本")
    print(f"\n輸出：{OUTPUT_SAMPLE}（{len(df):,} 列，{df['證券代碼'].nunique():,} 家）")
    print(f"      迴歸可用(核心+控制皆非空)：{reg_ok.sum():,} 列")
    print(f"      其中有 WATER_RATE：{(reg_ok & df['WATER_RATE'].notna()).sum():,} 列")
    print(f"      有揭露(WATER_DISC=1)：{(df['WATER_DISC']==1).sum():,} 列")
    print(f"\n輸出：{OUTPUT_STATS}（前 10 大產業）")
    with pd.option_context("display.unicode.east_asian_width", True,
                           "display.max_columns", None, "display.width", 200):
        print(stats.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
