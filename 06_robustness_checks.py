# -*- coding: utf-8 -*-
"""
06_robustness_checks.py
-----------------------
方向五「產業異質性」＋方向四「時間落差」交叉的子樣本檢定。

讀 04_model_data.csv（已含遞延交乘與高/低耗水產業分組 WaterUse）。
對每個 (水衡量 × 產業組 × 遞延期) 組合重跑成本黏性模型，比較核心係數 β3。

  水衡量：水回收率%（連續）、水揭露（虛擬）
  產業組：全樣本 / 高耗水 / 低耗水
  遞延期：t-0（當期）、t-1、t-2

方法：OLS + 產業(SASB)與年份固定效果 + 公司叢集穩健SE。

輸出：06_robustness_results.txt、06_robustness_coef.csv、06_robustness_summary.csv。
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm

INPUT_CSV = "04_model_data.csv"
OUT_TXT = "06_robustness_results.txt"
OUT_COEF = "06_robustness_coef.csv"
OUT_SUMMARY = "06_robustness_summary.csv"

CONTROL_INT = ["Size_D_dREV", "AI_D_dREV", "EI_D_dREV",
               "ROA_D_dREV", "Lev_D_dREV", "Decrease_D_dREV"]

# 各衡量在不同遞延期對應的（主效果欄, 三重交乘欄）
MEASURES = {
    # 水管理（原主題）
    "水回收率%": {0: ("Water_Rate", "WaterRate_D_dREV"),
                 1: ("WaterRate_l1", "WaterRate_D_dREV_l1"),
                 2: ("WaterRate_l2", "WaterRate_D_dREV_l2")},
    "水揭露": {0: ("Water_Disc", "WaterDisc_D_dREV"),
              1: ("WaterDisc_l1", "WaterDisc_D_dREV_l1"),
              2: ("WaterDisc_l2", "WaterDisc_D_dREV_l2")},
    # 廢棄物管理（延伸主題）
    "廢棄物密集度": {0: ("Waste_Intensity", "WasteInt_D_dREV"),
                   1: ("WasteInt_l1", "WasteInt_D_dREV_l1"),
                   2: ("WasteInt_l2", "WasteInt_D_dREV_l2")},
    "廢棄物揭露": {0: ("Waste_Disc", "WasteDisc_D_dREV"),
                 1: ("WasteDisc_l1", "WasteDisc_D_dREV_l1"),
                 2: ("WasteDisc_l2", "WasteDisc_D_dREV_l2")},
    "廢棄物裁罰金額": {0: ("Waste_Fine", "WasteFine_D_dREV"),
                    1: ("WasteFine_l1", "WasteFine_D_dREV_l1"),
                    2: ("WasteFine_l2", "WasteFine_D_dREV_l2")},
    "用水密集度": {0: ("Water_Intensity", "WaterInt_D_dREV"),
                 1: ("WaterInt_l1", "WaterInt_D_dREV_l1"),
                 2: ("WaterInt_l2", "WaterInt_D_dREV_l2")},
}
GROUPS = ["全樣本", "高耗水", "低耗水"]


def stars(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""


def run_spec(df, measure, group, lag, txt, coef, summary):
    water_main, water_triple = MEASURES[measure][lag]
    d = df.copy()
    if group == "高耗水":
        d = d[d["WaterUse"] == "高耗水"]
    elif group == "低耗水":
        d = d[d["WaterUse"] == "低耗水"]

    reg_vars = ["dLNREV", "D_x_dREV", water_triple, water_main] + CONTROL_INT
    d = d.dropna(subset=reg_vars + ["Y_dLNSGA", "Industry", "西元年份", "證券代碼"])
    d = d[d["Industry"].astype(str).str.strip() != ""]
    spec_name = f"{measure}｜{group}｜t-{lag}"

    if d["證券代碼"].nunique() < 5 or len(d) < 30 or d["Industry"].nunique() < 2:
        txt.append(f"[{spec_name}] 樣本不足(N={len(d)})，略過。")
        summary.append({"水衡量": measure, "產業組": group, "遞延期": f"t-{lag}",
                        "β2(D×ΔREV)": np.nan, "β2_sig": "—",
                        "β3(Water×D×ΔREV)": np.nan, "β3_sig": "樣本不足",
                        "N": len(d), "公司數": d["證券代碼"].nunique(), "adjR2": np.nan})
        return

    ind_d = pd.get_dummies(d["Industry"], prefix="Ind", drop_first=True)
    yr_d = pd.get_dummies(d["西元年份"].astype(int).astype(str), prefix="Yr", drop_first=True)
    X = pd.concat([d[reg_vars].reset_index(drop=True),
                   ind_d.reset_index(drop=True), yr_d.reset_index(drop=True)], axis=1)
    X = sm.add_constant(X, has_constant="add").astype(float)
    y = d["Y_dLNSGA"].reset_index(drop=True).astype(float)
    groups = d["證券代碼"].reset_index(drop=True)

    res = sm.OLS(y, X).fit(cov_type="cluster", cov_kwds={"groups": groups})
    n, k = int(res.nobs), d["證券代碼"].nunique()
    b2, p2 = res.params["D_x_dREV"], res.pvalues["D_x_dREV"]
    b3, p3 = res.params[water_triple], res.pvalues[water_triple]

    txt.append("=" * 72)
    txt.append(f"[{spec_name}]  N={n:,} 公司={k:,} adjR2={res.rsquared_adj:.4f}")
    for v in ["dLNREV", "D_x_dREV", water_triple, water_main]:
        b, se, t, p = res.params[v], res.bse[v], res.tvalues[v], res.pvalues[v]
        txt.append(f"  {v:<22}{b:>12.4f}{se:>11.4f}{t:>8.2f}{p:>9.4f} {stars(p)}")
        coef.append({"設定": spec_name, "變數": v, "係數": b, "穩健SE": se,
                     "t值": t, "p值": p, "顯著性": stars(p), "N": n})
    summary.append({"水衡量": measure, "產業組": group, "遞延期": f"t-{lag}",
                    "β2(D×ΔREV)": round(b2, 4), "β2_sig": stars(p2) or "n.s.",
                    "β3(Water×D×ΔREV)": round(b3, 4), "β3_sig": stars(p3) or "n.s.",
                    "N": n, "公司數": k, "adjR2": round(res.rsquared_adj, 4)})


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)

    txt, coef, summary = [], [], []
    txt.append("穩健性：產業異質性(高/低耗水) × 時間落差(t-0/1/2)\n")
    for measure in MEASURES:
        for group in GROUPS:
            for lag in (0, 1, 2):
                run_spec(df, measure, group, lag, txt, coef, summary)

    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write("\n".join(txt))
    pd.DataFrame(coef).to_csv(OUT_COEF, index=False, encoding="utf-8-sig")
    sdf = pd.DataFrame(summary)
    sdf.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")

    print("\n".join(txt[-40:]) if len(txt) > 40 else "\n".join(txt))
    print("===== 跨設定對照 (06_robustness_summary.csv) =====")
    with pd.option_context("display.unicode.east_asian_width", True,
                           "display.max_columns", None, "display.width", 240):
        print(sdf.to_string(index=False))
    print(f"\n輸出：{OUT_TXT}、{OUT_COEF}、{OUT_SUMMARY}")


if __name__ == "__main__":
    main()
