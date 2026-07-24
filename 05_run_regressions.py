# -*- coding: utf-8 -*-
"""
05_run_regressions.py
---------------------
成本黏性 (cost stickiness) 主迴歸：
  ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Water_Var×D×ΔLNREV)
           + β4·Water_Var + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε

方法：OLS（pooled）＋ 產業(SASB主產業)與年份固定效果(虛擬變數)
      ＋ 公司(證券代碼)叢集穩健標準誤。

模型 1（主分析）：Water_Var = Water_Rate（水回收率%）
模型 2（次分析）：Water_Var = Water_Disc（水揭露虛擬，做法 A）

係數解讀：
  β2 (D_x_dREV) < 0 且顯著 → 樣本存在成本黏性。
  β3 主分析 (WaterRate_D_dREV) < 0 → 水回收績效「加劇」黏性（H1）。
  β3 次分析 (WaterDisc_D_dREV) > 0 → 水資訊揭露「緩解」黏性（H2）。

輸出：05_regression_results.txt、05_regression_coef.csv。
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm

INPUT_CSV = "04_model_data.csv"
OUT_TXT = "05_regression_results.txt"
OUT_COEF = "05_regression_coef.csv"

CONTROL_INT = ["Size_D_dREV", "AI_D_dREV", "EI_D_dREV",
               "ROA_D_dREV", "Lev_D_dREV", "Decrease_D_dREV"]


def stars(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""


def run_model(df, water_main, water_triple, label, txt_lines, coef_rows):
    reg_vars = ["dLNREV", "D_x_dREV", water_triple, water_main] + CONTROL_INT
    needed = ["Y_dLNSGA"] + reg_vars + ["Industry", "Year", "證券代碼"]
    d = df.dropna(subset=[c for c in needed if c != "Industry"]).copy()
    d = d[d["Industry"].notna() & (d["Industry"].astype(str).str.strip() != "")]

    # 產業與年份固定效果（虛擬變數，drop_first 避免共線）
    ind_d = pd.get_dummies(d["Industry"], prefix="Industry", drop_first=True)
    yr_d = pd.get_dummies(d["Year"].astype(int).astype(str), prefix="Year", drop_first=True)

    X = pd.concat([d[reg_vars].reset_index(drop=True),
                   ind_d.reset_index(drop=True),
                   yr_d.reset_index(drop=True)], axis=1)
    X = sm.add_constant(X)
    X = X.astype(float)
    y = d["Y_dLNSGA"].reset_index(drop=True).astype(float)
    groups = d["證券代碼"].reset_index(drop=True)

    model = sm.OLS(y, X)
    res = model.fit(cov_type="cluster", cov_kwds={"groups": groups})

    n, k = int(res.nobs), d["證券代碼"].nunique()
    txt_lines.append("=" * 78)
    txt_lines.append(f"【{label}】 Water_Var = {water_main}")
    txt_lines.append(f"樣本 N = {n:,}；公司數 = {k:,}；調整後 R² = {res.rsquared_adj:.4f}")
    txt_lines.append("-" * 78)
    txt_lines.append(f"{'變數':<22}{'係數':>14}{'穩健SE':>12}{'t':>9}{'p':>9}  顯著")
    key_vars = ["dLNREV", "D_x_dREV", water_triple, water_main] + CONTROL_INT
    for v in key_vars:
        b, se, t, p = res.params[v], res.bse[v], res.tvalues[v], res.pvalues[v]
        s = stars(p)
        txt_lines.append(f"{v:<22}{b:>14.4f}{se:>12.4f}{t:>9.2f}{p:>9.4f}  {s}")
        coef_rows.append({"模型": label, "Water_Var": water_main, "變數": v,
                          "係數": b, "穩健SE": se, "t值": t, "p值": p,
                          "顯著性": s, "N": n, "公司數": k,
                          "adj_R2": res.rsquared_adj})
    txt_lines.append("（另含產業與年份固定效果，係數略）")

    # 解讀
    b2 = res.params["D_x_dREV"]; p2 = res.pvalues["D_x_dREV"]
    b3 = res.params[water_triple]; p3 = res.pvalues[water_triple]
    txt_lines.append("-" * 78)
    txt_lines.append(f"β2 (D×ΔLNREV) = {b2:.4f} ({stars(p2) or 'n.s.'})："
                     + ("存在成本黏性" if (b2 < 0 and p2 < 0.1) else "未顯著呈現黏性"))
    txt_lines.append(f"β3 ({water_triple}) = {b3:.4f} ({stars(p3) or 'n.s.'})：" + interp_b3(label, b3, p3))
    txt_lines.append("")
    return res


def interp_b3(label, b3, p3):
    if p3 >= 0.1:
        return "交乘不顯著"
    if "主分析" in label:
        return "水回收績效加劇黏性（支持H1）" if b3 < 0 else "水回收績效緩解黏性（與H1相反）"
    return "水資訊揭露緩解黏性（支持H2）" if b3 > 0 else "水資訊揭露加劇黏性（與H2相反）"


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)

    txt_lines, coef_rows = [], []
    txt_lines.append("成本黏性主迴歸結果（OLS + 產業/年份固定效果 + 公司叢集穩健SE）\n")
    run_model(df, "Water_Rate", "WaterRate_D_dREV", "模型1 主分析(水回收率%)", txt_lines, coef_rows)
    run_model(df, "Water_Disc", "WaterDisc_D_dREV", "模型2 次分析(水揭露)", txt_lines, coef_rows)

    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write("\n".join(txt_lines))
    pd.DataFrame(coef_rows).to_csv(OUT_COEF, index=False, encoding="utf-8-sig")

    print("\n".join(txt_lines))
    print(f"\n輸出：{OUT_TXT}、{OUT_COEF}")


if __name__ == "__main__":
    main()
