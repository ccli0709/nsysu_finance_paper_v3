# -*- coding: utf-8 -*-
"""
04_build_model_variables.py
---------------------------
從 03_sample_data.csv 建立成本黏性 (cost stickiness) 迴歸所需的交乘項與固定效果欄位。

主模型（study.md）：
  ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Water_Var×D×ΔLNREV)
           + β4·Water_Var + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε

- 連續變數採縮尾後 (_w) 版本降低極端值影響。
- 交乘欄名皆為安全 ASCII 名，另輸出 04_var_mapping.csv 對照原始意義。
- 產業(SASB主產業)與年份(西元年份)保留為欄位，交由 05 以固定效果(虛擬變數)處理。

輸出：04_model_data.csv、04_var_mapping.csv（皆 UTF-8-SIG）。
"""

import numpy as np
import pandas as pd

INPUT_CSV = "03_sample_data.csv"
OUTPUT_CSV = "04_model_data.csv"
OUTPUT_MAP = "04_var_mapping.csv"

# 模型欄位（優先取縮尾 _w 版本）
Y_COL = "dLNSGA_w"          # 應變數 ΔLNSGA
DREV_COL = "dLNREV_w"       # ΔLNREV
WATER_RATE_COL = "Water_Rate_w"  # 主分析水管理績效
CONTROL_COLS = {            # 安全名 -> 來源欄
    "Size": "Size_w",
    "AI": "AI_w",
    "EI": "EI_w",
    "ROA": "ROA_w",
    "Lev": "Lev_w",
    "Decrease": "Decrease",  # 虛擬變數不縮尾
}


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)

    out = pd.DataFrame()
    out["證券代碼"] = df["證券代碼"]
    out["西元年份"] = df["西元年份"].astype(int)
    out["Industry"] = df["SASB主產業"].astype(str)     # 產業固定效果
    out["Year"] = df["西元年份"].astype(int)            # 年份固定效果

    # 基本項
    out["Y_dLNSGA"] = pd.to_numeric(df[Y_COL], errors="coerce")
    out["dLNREV"] = pd.to_numeric(df[DREV_COL], errors="coerce")
    out["D"] = pd.to_numeric(df["D"], errors="coerce")
    out["Water_Rate"] = pd.to_numeric(df[WATER_RATE_COL], errors="coerce")
    out["Water_Disc"] = pd.to_numeric(df["Water_Disc"], errors="coerce")
    out["Water_Disc_GRI"] = pd.to_numeric(df["Water_Disc_GRI"], errors="coerce")

    # 控制變數（主效果）
    for safe, src in CONTROL_COLS.items():
        out[safe] = pd.to_numeric(df[src], errors="coerce")

    # 核心交乘
    out["D_x_dREV"] = out["D"] * out["dLNREV"]                       # β2：成本黏性
    out["WaterRate_D_dREV"] = out["Water_Rate"] * out["D_x_dREV"]   # β3：水績效×黏性
    out["WaterDisc_D_dREV"] = out["Water_Disc"] * out["D_x_dREV"]   # β3：水揭露×黏性
    out["WaterDiscGRI_D_dREV"] = out["Water_Disc_GRI"] * out["D_x_dREV"]  # 穩健性

    # 控制變數 × (D×ΔLNREV)（ABJ 慣例）
    for safe in CONTROL_COLS:
        out[f"{safe}_D_dREV"] = out[safe] * out["D_x_dREV"]

    # ---- 方向四：遞延期水管理主效果與三重交乘（t-1、t-2）----
    LAG_MAP = {
        "WaterRate": "Water_Rate_w",       # 水回收率%（縮尾）
        "WaterDisc": "Water_Disc",         # 水揭露做法A
        "WaterDiscGRI": "Water_Disc_GRI",  # 水揭露做法B
        "WaterRateProc": "製程水回收率%_w", # 製程水回收率%（縮尾）
    }
    for base, src in LAG_MAP.items():
        for k in (1, 2):
            col = f"{src}_l{k}"
            if col in df.columns:
                out[f"{base}_l{k}"] = pd.to_numeric(df[col], errors="coerce")
                out[f"{base}_D_dREV_l{k}"] = out[f"{base}_l{k}"] * out["D_x_dREV"]

    # ---- 方向五：高/低耗水產業分類（依 TEJ Company DB 之 TSE 產業別，精準定義）----
    # 依水資源相關文獻與台灣製造業實務，界定高耗水（用水密集）產業之 TSE 代碼。
    out["TSE代碼"] = df["TSE產業_代碼"].astype(str).str.strip()
    out["TSE產業"] = df["TSE產業_名稱"].astype(str).str.strip()
    out.loc[out["TSE代碼"].isin(["nan", ""]), "TSE代碼"] = np.nan
    out.loc[out["TSE產業"].isin(["nan", ""]), "TSE產業"] = np.nan

    HIGH_WATER_TSE = {
        "M2324": "半導體業", "M2326": "光電業", "M2328": "電子零組件",
        "M1400": "紡織纖維", "M1900": "造紙工業", "M2000": "鋼鐵工業",
        "M1721": "化學工業", "M1100": "水泥工業", "M1800": "玻璃陶瓷",
        "M1200": "食品工業", "M1300": "塑膠工業", "M2100": "橡膠工業",
        "M9700": "油電燃氣業",
    }
    out["WaterUse"] = pd.Series(np.nan, index=out.index, dtype="object")
    has_tse = out["TSE代碼"].notna()
    out.loc[has_tse, "WaterUse"] = "低耗水"
    out.loc[out["TSE代碼"].isin(HIGH_WATER_TSE), "WaterUse"] = "高耗水"

    out = out.sort_values(["證券代碼", "西元年份"]).reset_index(drop=True)
    out.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    # 產業分組對照輸出（供論文與檢視）：各 TSE 產業之觀測數與分組
    dens_df = (out.dropna(subset=["TSE產業"])
               .groupby(["TSE代碼", "TSE產業"])
               .agg(觀測數=("證券代碼", "size"),
                    公司數=("證券代碼", "nunique"),
                    有水資料=("Water_Rate", lambda s: int(s.notna().sum())))
               .reset_index())
    dens_df["耗水分組"] = np.where(dens_df["TSE代碼"].isin(HIGH_WATER_TSE), "高耗水", "低耗水")
    dens_df = dens_df.sort_values(["耗水分組", "觀測數"], ascending=[True, False])

    # 變數對照表
    mapping = [
        ("Y_dLNSGA", "ΔLNSGA = LN(營業費用_t) − LN(營業費用_t-1)（縮尾）", "應變數"),
        ("dLNREV", "ΔLNREV = LN(營收_t) − LN(營收_t-1)（縮尾）", "自變數"),
        ("D", "收入下降虛擬（營收_t < 營收_t-1 = 1）", "自變數"),
        ("D_x_dREV", "D × ΔLNREV（β2：整體成本黏性，預期負）", "核心交乘"),
        ("Water_Rate", "水回收率%（縮尾，主分析水績效）", "自變數"),
        ("Water_Disc", "水揭露虛擬-做法A（有水量紀錄=1）", "自變數"),
        ("Water_Disc_GRI", "水揭露虛擬-做法B（GRI揭露度>0=1）", "自變數(穩健)"),
        ("WaterRate_D_dREV", "Water_Rate × D × ΔLNREV（β3：H1）", "三重交乘"),
        ("WaterDisc_D_dREV", "Water_Disc × D × ΔLNREV（β3：H2）", "三重交乘"),
        ("WaterDiscGRI_D_dREV", "Water_Disc_GRI × D × ΔLNREV（穩健）", "三重交乘"),
        ("Size", "LN(資產總額)（縮尾）", "控制"),
        ("AI", "資產密集度=資產總額/營收（縮尾）", "控制"),
        ("EI", "員工密集度=員工人數/營收（縮尾）", "控制"),
        ("ROA", "資產報酬率 ROA(A)稅後息前（縮尾）", "控制"),
        ("Lev", "財務槓桿=負債比率（縮尾）", "控制"),
        ("Decrease", "連續兩年營收下降虛擬", "控制"),
        ("Size_D_dREV", "Size × D × ΔLNREV", "控制交乘"),
        ("AI_D_dREV", "AI × D × ΔLNREV", "控制交乘"),
        ("EI_D_dREV", "EI × D × ΔLNREV", "控制交乘"),
        ("ROA_D_dREV", "ROA × D × ΔLNREV", "控制交乘"),
        ("Lev_D_dREV", "Lev × D × ΔLNREV", "控制交乘"),
        ("Decrease_D_dREV", "Decrease × D × ΔLNREV", "控制交乘"),
        ("WaterRate_D_dREV_l1", "Water_Rate(t-1) × D × ΔLNREV（遞延一期，方向四）", "遞延交乘"),
        ("WaterRate_D_dREV_l2", "Water_Rate(t-2) × D × ΔLNREV（遞延兩期，方向四）", "遞延交乘"),
        ("WaterDisc_D_dREV_l1", "Water_Disc(t-1) × D × ΔLNREV（遞延一期）", "遞延交乘"),
        ("WaterDisc_D_dREV_l2", "Water_Disc(t-2) × D × ΔLNREV（遞延兩期）", "遞延交乘"),
        ("WaterUse", "耗水分組（高/低耗水產業，方向五）", "分組"),
        ("Industry", "SASB主產業（產業固定效果）", "固定效果"),
        ("Year", "西元年份（年份固定效果）", "固定效果"),
    ]
    pd.DataFrame(mapping, columns=["安全欄名", "定義", "角色"]).to_csv(
        OUTPUT_MAP, index=False, encoding="utf-8-sig")
    dens_df.to_csv("04_industry_wateruse.csv", index=False, encoding="utf-8-sig")

    # 摘要
    print("===== 建模變數摘要 =====")
    print(f"輸出：{OUTPUT_CSV}（{len(out):,} 列，{out['證券代碼'].nunique():,} 家）")
    print(f"輸出：{OUTPUT_MAP}（{len(mapping)} 個變數對照）")
    base = ["Y_dLNSGA", "dLNREV", "D", "D_x_dREV"] + list(CONTROL_COLS)
    m1_ok = out[base].notna().all(axis=1)
    print(f"\n模型1(主分析 Water_Rate) 可用列："
          f"{(m1_ok & out['WaterRate_D_dREV'].notna()).sum():,}")
    print(f"模型2(次分析 Water_Disc) 可用列：{(m1_ok & out['WaterDisc_D_dREV'].notna()).sum():,}")
    print(f"產業數：{out['Industry'].nunique()}；年份：{out['Year'].min()}-{out['Year'].max()}")
    print("\n遞延期可用列（核心+控制皆非空）：")
    for c in ["WaterRate_D_dREV", "WaterRate_D_dREV_l1", "WaterRate_D_dREV_l2",
              "WaterDisc_D_dREV", "WaterDisc_D_dREV_l1", "WaterDisc_D_dREV_l2"]:
        if c in out.columns:
            print(f"  {c:<22}: {(m1_ok & out[c].notna()).sum():,}")
    print("\n高/低耗水產業分組：")
    print(dens_df.to_string(index=False))


if __name__ == "__main__":
    main()
