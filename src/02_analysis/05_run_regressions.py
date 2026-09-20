# -*- coding: utf-8 -*-
"""
05_run_regressions.py
---------------------
成本黏性 (cost stickiness) 主迴歸 + 方向四「時間落差效應 (Time Lag)」。

模型：
  ΔLNSGA = β0 + β1·ΔLNREV + β2·(D×ΔLNREV) + β3·(Water_Var(t-k)×D×ΔLNREV)
           + β4·Water_Var(t-k) + Σ Controls×(D×ΔLNREV) + Σ Industry + Σ Year + ε

- 水管理變數不新增，只做期數平移：k=0（當期）、1（t-1）、2（t-2）。
- 模型1（主分析）：Water_Var = 水回收率%（Water_Rate）
- 模型2（次分析）：Water_Var = 水揭露虛擬（Water_Disc）
- 方法：OLS + 產業/年份固定效果 + 公司叢集穩健標準誤。

理論預期：水管理投資效益需時間發酵，遞延項（L1/L2）之 β3 可能較當期顯著。

輸出：05_regression_results.txt、05_regression_coef.csv。
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm

INPUT_CSV = "data/processed/04_model_data.csv"
OUT_TXT = "results/05_regression_results.txt"
OUT_COEF = "results/05_regression_coef.csv"

CONTROL_INT = ["Size_D_dREV", "AI_D_dREV", "EI_D_dREV",
               "ROA_D_dREV", "Lev_D_dREV", "Decrease_D_dREV"]


def stars(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""


def interp_b3(kind, b3, p3):
    if p3 >= 0.1:
        return "交乘不顯著"
    neg = "加劇黏性" if b3 < 0 else "緩解黏性"
    hyp = {
        "主分析": ("水回收率（補充：實質投入）", "補充檢驗", "負"),
        "次分析": ("水資訊揭露（補充：形式揭露）", "補充檢驗", "正"),
        "廢棄物主分析": ("廢棄物密集度", "核心 H1b", "負"),
        "廢棄物次分析": ("廢棄物揭露品質（補充：形式揭露）", "補充檢驗", "正"),
        "廢棄物穩健": ("廢棄物違規裁罰（補充：法遵裁罰）", "補充檢驗", "負"),
        "用水密集度": ("用水密集度", "核心 H1a", "負"),
    }
    name, h, expect = hyp.get(kind, (kind, "", ""))
    support = "支持" if ((expect == "負" and b3 < 0) or (expect == "正" and b3 > 0)) else "與假設相反"
    return f"{name}{neg}（{support}{h}）"


def run_model(df, water_main, water_triple, label, kind, lag, txt_lines, coef_rows):
    reg_vars = ["dLNREV", "D_x_dREV", water_triple, water_main] + CONTROL_INT
    d = df.dropna(subset=reg_vars + ["Y_dLNSGA", "西元年份", "證券代碼"]).copy()
    d = d[d["Industry"].notna() & (d["Industry"].astype(str).str.strip() != "")]
    if d["證券代碼"].nunique() < 5 or len(d) < 30:
        txt_lines.append(f"[{label}] 樣本不足({len(d)})，略過。\n")
        return

    ind_d = pd.get_dummies(d["Industry"], prefix="Industry", drop_first=True)
    yr_d = pd.get_dummies(d["Year"].astype(int).astype(str), prefix="Year", drop_first=True)
    X = pd.concat([d[reg_vars].reset_index(drop=True),
                   ind_d.reset_index(drop=True),
                   yr_d.reset_index(drop=True)], axis=1)
    X = sm.add_constant(X, has_constant="add").astype(float)
    y = d["Y_dLNSGA"].reset_index(drop=True).astype(float)
    groups = d["證券代碼"].reset_index(drop=True)

    res = sm.OLS(y, X).fit(cov_type="cluster", cov_kwds={"groups": groups})
    n, k = int(res.nobs), d["證券代碼"].nunique()

    txt_lines.append("=" * 78)
    txt_lines.append(f"【{label}】 Water_Var = {water_main}（遞延期 t-{lag}）")
    txt_lines.append(f"樣本 N = {n:,}；公司數 = {k:,}；調整後 R² = {res.rsquared_adj:.4f}")
    txt_lines.append("-" * 78)
    txt_lines.append(f"{'變數':<24}{'係數':>14}{'穩健SE':>12}{'t':>9}{'p':>9}  顯著")
    for v in ["dLNREV", "D_x_dREV", water_triple, water_main] + CONTROL_INT:
        b, se, t, p = res.params[v], res.bse[v], res.tvalues[v], res.pvalues[v]
        txt_lines.append(f"{v:<24}{b:>14.4f}{se:>12.4f}{t:>9.2f}{p:>9.4f}  {stars(p)}")
        coef_rows.append({"模型": label, "類型": kind, "遞延期": lag,
                          "Water_Var": water_main, "變數": v, "角色": role_of(v, water_triple, water_main),
                          "係數": b, "穩健SE": se, "t值": t, "p值": p,
                          "顯著性": stars(p), "N": n, "公司數": k,
                          "adj_R2": res.rsquared_adj})
    b2, p2 = res.params["D_x_dREV"], res.pvalues["D_x_dREV"]
    b3, p3 = res.params[water_triple], res.pvalues[water_triple]
    txt_lines.append("-" * 78)
    txt_lines.append(f"β2 (D×ΔLNREV) = {b2:.4f} ({stars(p2) or 'n.s.'})："
                     + ("存在成本黏性" if (b2 < 0 and p2 < 0.1) else "未顯著呈現黏性"))
    txt_lines.append(f"β3 ({water_triple}) = {b3:.4f} ({stars(p3) or 'n.s.'})：" + interp_b3(kind, b3, p3))
    txt_lines.append("")


def role_of(v, triple, main):
    if v == triple:
        return "β3(三重交乘)"
    if v == main:
        return "β4(水主效果)"
    if v == "D_x_dREV":
        return "β2(黏性)"
    if v == "dLNREV":
        return "β1"
    return "控制交乘"


def main():
    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    df["證券代碼"] = df["證券代碼"].astype(str)

    specs = [
        # ---- 水管理（原主題）----
        ("模型1 水回收率% L0(當期)", "Water_Rate", "WaterRate_D_dREV", "主分析", 0),
        ("模型1 水回收率% L1(t-1)", "WaterRate_l1", "WaterRate_D_dREV_l1", "主分析", 1),
        ("模型1 水回收率% L2(t-2)", "WaterRate_l2", "WaterRate_D_dREV_l2", "主分析", 2),
        ("模型2 水揭露 L0(當期)", "Water_Disc", "WaterDisc_D_dREV", "次分析", 0),
        ("模型2 水揭露 L1(t-1)", "WaterDisc_l1", "WaterDisc_D_dREV_l1", "次分析", 1),
        ("模型2 水揭露 L2(t-2)", "WaterDisc_l2", "WaterDisc_D_dREV_l2", "次分析", 2),
        # ---- 廢棄物管理（延伸主題）----
        ("模型3 廢棄物密集度 L0(當期)", "Waste_Intensity", "WasteInt_D_dREV", "廢棄物主分析", 0),
        ("模型3 廢棄物密集度 L1(t-1)", "WasteInt_l1", "WasteInt_D_dREV_l1", "廢棄物主分析", 1),
        ("模型3 廢棄物密集度 L2(t-2)", "WasteInt_l2", "WasteInt_D_dREV_l2", "廢棄物主分析", 2),
        ("模型4 廢棄物揭露 L0(當期)", "Waste_Disc", "WasteDisc_D_dREV", "廢棄物次分析", 0),
        ("模型4 廢棄物揭露 L1(t-1)", "WasteDisc_l1", "WasteDisc_D_dREV_l1", "廢棄物次分析", 1),
        ("模型4 廢棄物揭露 L2(t-2)", "WasteDisc_l2", "WasteDisc_D_dREV_l2", "廢棄物次分析", 2),
        ("模型5 廢棄物裁罰金額 L0(當期)", "Waste_Fine", "WasteFine_D_dREV", "廢棄物穩健", 0),
        ("模型5 廢棄物裁罰金額 L1(t-1)", "WasteFine_l1", "WasteFine_D_dREV_l1", "廢棄物穩健", 1),
        ("模型5 廢棄物裁罰金額 L2(t-2)", "WasteFine_l2", "WasteFine_D_dREV_l2", "廢棄物穩健", 2),
        # ---- 用水密集度（升為主模型，使用密集度核心）----
        ("模型6 用水密集度 L0(當期)", "Water_Intensity", "WaterInt_D_dREV", "用水密集度", 0),
        ("模型6 用水密集度 L1(t-1)", "WaterInt_l1", "WaterInt_D_dREV_l1", "用水密集度", 1),
        ("模型6 用水密集度 L2(t-2)", "WaterInt_l2", "WaterInt_D_dREV_l2", "用水密集度", 2),
    ]

    txt_lines, coef_rows = [], []
    txt_lines.append("成本黏性主迴歸 + 時間落差效應（OLS + 產業/年份FE + 公司叢集SE）\n")
    for label, wm, wt, kind, lag in specs:
        run_model(df, wm, wt, label, kind, lag, txt_lines, coef_rows)

    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write("\n".join(txt_lines))
    pd.DataFrame(coef_rows).to_csv(OUT_COEF, index=False, encoding="utf-8-sig")

    print("\n".join(txt_lines))
    # 遞延期 β3 對照
    cdf = pd.DataFrame(coef_rows)
    b3 = cdf[cdf["角色"] == "β3(三重交乘)"][["模型", "類型", "遞延期", "係數", "顯著性", "N"]]
    print("===== 遞延期 β3（Water×D×ΔLNREV）對照 =====")
    with pd.option_context("display.unicode.east_asian_width", True, "display.width", 200):
        print(b3.to_string(index=False))
    print(f"\n輸出：{OUT_TXT}、{OUT_COEF}")


if __name__ == "__main__":
    main()
