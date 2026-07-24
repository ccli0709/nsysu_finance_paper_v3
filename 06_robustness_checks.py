# -*- coding: utf-8 -*-
"""
06_robustness_checks.py
-----------------------
成本黏性主結論之穩健性檢定。輸入回到 03_sample_data.csv（全製造業），
逐一替換衡量方式並重跑主模型（OLS + 產業/年份固定效果 + 公司叢集穩健SE）。

四個維度：
  1. 水績效衡量替換：水回收率% ↔ 製程水回收率%
  2. 揭露定義替換：做法A（有水量紀錄）↔ 做法B（GRI 揭露度>0）
  3. 營業費用(SG&A)定義替換：營業費用 ↔ 推銷費用+管理費用
  4. 產業樣本維度：全製造業 ↔ 限定水資料充足產業

輸出：06_robustness_results.txt、06_robustness_coef.csv、06_robustness_summary.csv。
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm

INPUT_CSV = "03_sample_data.csv"
OUT_TXT = "06_robustness_results.txt"
OUT_COEF = "06_robustness_coef.csv"
OUT_SUMMARY = "06_robustness_summary.csv"

# 水資料充足產業（依 03_industry_stats：有水資料筆數 ≥ 80）
WATER_RICH = ["科技與通訊", "資源轉化", "消費品"]

CONTROL_SRC = {  # 安全名 -> 03 檔來源欄（皆縮尾/虛擬）
    "Size": "Size_w", "AI": "AI_w", "EI": "EI_w",
    "ROA": "ROA_w", "Lev": "Lev_w", "Decrease": "Decrease",
}
CONTROL_INT = [f"{k}_D_dREV" for k in CONTROL_SRC]


def stars(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""


def run_spec(df, y_src, water_src, water_kind, industries, spec_name,
             txt_lines, coef_rows, summary_rows, dims):
    d = df.copy()
    if industries is not None:
        d = d[d["SASB主產業"].isin(industries)]

    d = d.assign(
        Y=pd.to_numeric(d[y_src], errors="coerce"),
        dREV=pd.to_numeric(d["dLNREV_w"], errors="coerce"),
        D=pd.to_numeric(d["D"], errors="coerce"),
        Wv=pd.to_numeric(d[water_src], errors="coerce"),
    )
    d["D_x_dREV"] = d["D"] * d["dREV"]
    d["W_D_dREV"] = d["Wv"] * d["D_x_dREV"]
    for k, src in CONTROL_SRC.items():
        d[k] = pd.to_numeric(d[src], errors="coerce")
        d[f"{k}_D_dREV"] = d[k] * d["D_x_dREV"]

    reg_vars = ["dREV", "D_x_dREV", "W_D_dREV", "Wv"] + CONTROL_INT
    d = d.dropna(subset=["Y"] + reg_vars + ["SASB主產業", "西元年份", "證券代碼"])
    d = d[d["SASB主產業"].astype(str).str.strip() != ""]
    if d["證券代碼"].nunique() < 5 or len(d) < 30:
        txt_lines.append(f"[{spec_name}] 樣本不足，略過。\n")
        return

    ind_d = pd.get_dummies(d["SASB主產業"], prefix="Ind", drop_first=True)
    yr_d = pd.get_dummies(d["西元年份"].astype(int).astype(str), prefix="Yr", drop_first=True)
    X = pd.concat([d[reg_vars].reset_index(drop=True),
                   ind_d.reset_index(drop=True),
                   yr_d.reset_index(drop=True)], axis=1)
    X = sm.add_constant(X, has_constant="add").astype(float)
    y = d["Y"].reset_index(drop=True).astype(float)
    groups = d["證券代碼"].reset_index(drop=True)

    res = sm.OLS(y, X).fit(cov_type="cluster", cov_kwds={"groups": groups})
    n, k = int(res.nobs), d["證券代碼"].nunique()

    b2, p2 = res.params["D_x_dREV"], res.pvalues["D_x_dREV"]
    b3, p3 = res.params["W_D_dREV"], res.pvalues["W_D_dREV"]

    txt_lines.append("=" * 74)
    txt_lines.append(f"[{spec_name}]  N={n:,}  公司={k:,}  adjR2={res.rsquared_adj:.4f}")
    for v in ["dREV", "D_x_dREV", "W_D_dREV", "Wv"] + CONTROL_INT:
        b, se, t, p = res.params[v], res.bse[v], res.tvalues[v], res.pvalues[v]
        txt_lines.append(f"  {v:<16}{b:>12.4f}{se:>11.4f}{t:>8.2f}{p:>9.4f} {stars(p)}")
        coef_rows.append({"設定": spec_name, "變數": v, "係數": b, "穩健SE": se,
                          "t值": t, "p值": p, "顯著性": stars(p), "N": n})
    txt_lines.append("")

    summary_rows.append({
        "設定": spec_name, **dims,
        "β2(D×ΔREV)": round(b2, 4), "β2_sig": stars(p2) or "n.s.",
        "β3(Water×D×ΔREV)": round(b3, 4), "β3_sig": stars(p3) or "n.s.",
        "水衡量型態": water_kind, "N": n, "公司數": k,
        "adjR2": round(res.rsquared_adj, 4),
    })


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)

    txt, coef, summary = [], [], []
    txt.append("成本黏性穩健性檢定（OLS + 產業/年份FE + 公司叢集SE）\n")

    # 基準
    run_spec(df, "dLNSGA_w", "水回收率%_w", "水回收率%(連續)", None,
             "B0 基準：水回收率%×全產業×營業費用", txt, coef, summary,
             {"維度": "基準", "SGA": "營業費用", "產業": "全製造業"})

    # 維度1：水績效衡量替換
    run_spec(df, "dLNSGA_w", "製程水回收率%_w", "製程水回收率%(連續)", None,
             "R1 水衡量：製程水回收率%", txt, coef, summary,
             {"維度": "水衡量替換", "SGA": "營業費用", "產業": "全製造業"})

    # 維度2：揭露定義替換
    run_spec(df, "dLNSGA_w", "Water_Disc", "揭露做法A(虛擬)", None,
             "R2 揭露A：有水量紀錄", txt, coef, summary,
             {"維度": "揭露定義", "SGA": "營業費用", "產業": "全製造業"})
    run_spec(df, "dLNSGA_w", "Water_Disc_GRI", "揭露做法B(虛擬)", None,
             "R3 揭露B：GRI揭露度>0", txt, coef, summary,
             {"維度": "揭露定義", "SGA": "營業費用", "產業": "全製造業"})

    # 維度3：SG&A 定義替換
    run_spec(df, "dLNSGA_alt_w", "水回收率%_w", "水回收率%(連續)", None,
             "R4 SG&A：推銷+管理費用", txt, coef, summary,
             {"維度": "SG&A定義", "SGA": "推銷+管理", "產業": "全製造業"})

    # 維度4：產業樣本維度（限水資料充足產業）
    run_spec(df, "dLNSGA_w", "水回收率%_w", "水回收率%(連續)", WATER_RICH,
             "R5 產業：限水資料充足產業(水回收率%)", txt, coef, summary,
             {"維度": "產業樣本", "SGA": "營業費用", "產業": "水資料充足"})
    run_spec(df, "dLNSGA_w", "Water_Disc", "揭露做法A(虛擬)", WATER_RICH,
             "R6 產業：限水資料充足產業(揭露A)", txt, coef, summary,
             {"維度": "產業樣本", "SGA": "營業費用", "產業": "水資料充足"})

    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write("\n".join(txt))
    pd.DataFrame(coef).to_csv(OUT_COEF, index=False, encoding="utf-8-sig")
    sm_df = pd.DataFrame(summary)
    sm_df.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")

    print("\n".join(txt))
    print("===== 跨設定對照 (06_robustness_summary.csv) =====")
    with pd.option_context("display.unicode.east_asian_width", True,
                           "display.max_columns", None, "display.width", 220):
        print(sm_df.to_string(index=False))
    print(f"\n輸出：{OUT_TXT}、{OUT_COEF}、{OUT_SUMMARY}")


if __name__ == "__main__":
    main()
