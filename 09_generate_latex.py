# -*- coding: utf-8 -*-
"""
09_generate_latex.py
--------------------
從 04_model_data.csv、05_regression_coef.csv 及 06_robustness_summary.csv 讀取數據，
生成符合標準 LaTeX (XeLaTeX / ctex) 格式的完整學術論文原始碼：main.tex。

字型設定規格：
- 中文字型：標楷體 (BiauKai / DFKai-SB)
- 英文字型：Times New Roman
"""

import os
import numpy as np
import pandas as pd

MODEL_CSV = "04_model_data.csv"
COEF_CSV = "05_regression_coef.csv"
ROB_CSV = "06_robustness_summary.csv"
OUT_TEX = "main.tex"

# 載入資料
df = pd.read_csv(MODEL_CSV, encoding="utf-8-sig", low_memory=False)
coef = pd.read_csv(COEF_CSV, encoding="utf-8-sig")
rob = pd.read_csv(ROB_CSV, encoding="utf-8-sig")

N_obs = len(df)
N_firm = df["證券代碼"].nunique()

def tex_escape(text):
    """轉義 LaTeX 特殊字元"""
    if not isinstance(text, str):
        text = str(text)
    replacements = [
        ('\\', r'\textbackslash{}'),
        ('&', r'\&'),
        ('%', r'\%'),
        ('$', r'\$'),
        ('#', r'\#'),
        ('_', r'\_'),
        ('{', r'\{'),
        ('}', r'\}'),
        ('~', r'\textasciitilde{}'),
        ('^', r'\textasciicircum{}'),
    ]
    for old, new in replacements:
        text = text.replace(old, new)
    return text

# 變數說明定義
DESC_VARS = [
    ("Y_dLNSGA", r"$\Delta\text{LNSGA}$", "營業費用對數變動率"),
    ("dLNREV", r"$\Delta\text{LNREV}$", "營業收入對數變動率"),
    ("D", "$D$", "收入下降虛擬變數"),
    ("Water_Intensity", r"\text{Water\_Intensity}", "用水密集度（千公噸/千元營收）"),
    ("Waste_Intensity", r"\text{Waste\_Intensity}", "廢棄物密集度（公噸/百萬營收）"),
    ("Water_Rate", r"\text{Water\_Rate}", r"水回收率\%"),
    ("Waste_Disc", r"\text{Waste\_Disc}", "GRI廢棄物揭露度"),
    ("Waste_Fine", r"\text{Waste\_Fine}", "廢棄物裁罰金額/資產"),
    ("Size", r"\text{Size}", "公司規模 LN(資產總額)"),
    ("AI", r"\text{AI}", "資產密集度 (資產總額/營收)"),
    ("EI", r"\text{EI}", "員工密集度 (員工人數/營收)"),
    ("ROA", r"\text{ROA}", r"資產報酬率 (\%)"),
    ("Lev", r"\text{Lev}", r"負債比率 (\%)"),
    ("Decrease", r"\text{Decrease}", "連續兩年營收下降虛擬變數")
]

CORR_VARS = [
    ("Y_dLNSGA", r"$\Delta\text{LNSGA}$"),
    ("dLNREV", r"$\Delta\text{LNREV}$"),
    ("Water_Intensity", r"\text{Water\_Int}"),
    ("Waste_Intensity", r"\text{Waste\_Int}"),
    ("Water_Rate", r"\text{Water\_Rate}"),
    ("Size", r"\text{Size}"),
    ("AI", r"\text{AI}"),
    ("ROA", r"\text{ROA}"),
    ("Lev", r"\text{Lev}")
]

L0_MODELS = [
    ("模型1 水回收率% L0(當期)", "模型 1\\\\(水回收率\\%)"),
    ("模型2 水揭露 L0(當期)", "模型 2\\\\(水揭露)"),
    ("模型3 廢棄物密集度 L0(當期)", "模型 3\\\\(廢棄物密集度)"),
    ("模型4 廢棄物揭露 L0(當期)", "模型 4\\\\(廢棄物揭露)"),
    ("模型5 廢棄物裁罰金額 L0(當期)", "模型 5\\\\(廢棄物裁罰)"),
    ("模型6 用水密集度 L0(當期)", "模型 6\\\\(用水密集度)"),
]

ROLE_ROWS = [
    (r"$\beta_1$", r"$\Delta\text{LNREV}$", "dLNREV"),
    (r"$\beta_2$", r"$D \times \Delta\text{LNREV}$", "D_x_dREV"),
    (r"$\beta_3$", r"$\text{Env\_Var} \times D \times \Delta\text{LNREV}$", None),
    (r"$\beta_4$", r"$\text{Env\_Var}$", None),
    ("", r"$\text{Size} \times D \times \Delta\text{LNREV}$", "Size_D_dREV"),
    ("", r"$\text{AI} \times D \times \Delta\text{LNREV}$", "AI_D_dREV"),
    ("", r"$\text{EI} \times D \times \Delta\text{LNREV}$", "EI_D_dREV"),
    ("", r"$\text{ROA} \times D \times \Delta\text{LNREV}$", "ROA_D_dREV"),
    ("", r"$\text{Lev} \times D \times \Delta\text{LNREV}$", "Lev_D_dREV"),
    ("", r"$\text{Decrease} \times D \times \Delta\text{LNREV}$", "Decrease_D_dREV"),
]

def build_desc_table_tex():
    lines = []
    lines.append(r"\begin{table}[H]")
    lines.append(r"\centering")
    lines.append(r"\caption{主要變數敘述統計表}")
    lines.append(r"\label{tab:desc_stats}")
    lines.append(r"\small")
    lines.append(r"\begin{tabular}{lrrrrrrrr}")
    lines.append(r"\toprule")
    lines.append(r"變數名稱 & 樣本數(N) & 平均值 & 標準差 & 最小值 & P25 & 中位數 & P75 & 最大值 \\")
    lines.append(r"\midrule")
    for col, lab, desc in DESC_VARS:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(s) == 0:
            continue
        line = f"{lab} & {len(s):,} & {s.mean():.3f} & {s.std():.3f} & {s.min():.3f} & {s.quantile(.25):.3f} & {s.median():.3f} & {s.quantile(.75):.3f} & {s.max():.3f} \\\\"
        lines.append(line)
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：本表樣本取自台灣上市櫃製造業（2014--2024）。所有連續型變數均已進行上下 1\% 之縮尾處理（Winsorization）。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_corr_table_tex():
    cols = [c for c, _ in CORR_VARS]
    labs = [l for _, l in CORR_VARS]
    sub = df[cols].apply(pd.to_numeric, errors="coerce")
    cm = sub.corr()
    n_sample = len(sub.dropna())
    
    lines = []
    lines.append(r"\begin{table}[H]")
    lines.append(r"\centering")
    lines.append(r"\caption{主要變數 Pearson 相關係數矩陣}")
    lines.append(r"\label{tab:corr_matrix}")
    lines.append(r"\footnotesize")
    col_spec = "l" + "c" * len(cols)
    lines.append(f"\\begin{{{col_spec}}}")
    lines.append(r"\toprule")
    header = "變數" + "".join([f" & ({i+1})" for i in range(len(cols))]) + r" \\"
    lines.append(header)
    lines.append(r"\midrule")
    
    for i, lab in enumerate(labs):
        row_str = f"({i+1}) {lab}"
        for j in range(len(cols)):
            if j > i:
                row_str += " & "
            elif i == j:
                row_str += " & 1.000"
            else:
                r_val = cm.iloc[i, j]
                t_val = abs(r_val) * np.sqrt(max(n_sample - 2, 1) / max(1e-9, 1 - r_val * r_val))
                if t_val >= 1.96:
                    row_str += f" & \\textbf{{{r_val:.3f}}}"
                else:
                    row_str += f" & {r_val:.3f}"
        row_str += r" \\"
        lines.append(row_str)
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：對角線下方為 Pearson 相關係數；粗體字代表在 5\% 水準下達顯著（$|t| \ge 1.96$）。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def get_coef_cell(model_name, role=None, var=None):
    r = coef[coef["模型"] == model_name]
    if role:
        rr = r[r["角色"] == role]
    else:
        rr = r[r["變數"] == var]
    if rr.empty:
        return "", ""
    b = rr["係數"].iloc[0]
    s = rr["顯著性"].iloc[0]
    t = rr["t值"].iloc[0]
    s_str = "" if pd.isna(s) else str(s)
    b_str = f"{b:.4f}{s_str}"
    t_str = f"({t:.2f})"
    return b_str, t_str

def build_reg_main_table_tex():
    lines = []
    lines.append(r"\begin{table}[H]")
    lines.append(r"\centering")
    lines.append(r"\caption{成本黏性主迴歸結果（當期，模型 1 至模型 6）}")
    lines.append(r"\label{tab:reg_main}")
    lines.append(r"\scriptsize")
    lines.append(r"\begin{tabular}{lcccccc}")
    lines.append(r"\toprule")
    head_line = "變數名稱 & " + " & ".join([lab for _, lab in L0_MODELS]) + r" \\"
    lines.append(head_line)
    lines.append(r"\midrule")
    
    for greek, label, var in ROLE_ROWS:
        role = {"$\beta_1$": "β1", r"$\beta_2$": "β2(黏性)", r"$\beta_3$": "β3(三重交乘)", r"$\beta_4$": "β4(水主效果)"}.get(greek)
        disp = f"{label} ({greek})" if greek else f"{label}"
        
        b_cells = []
        t_cells = []
        for m_name, _ in L0_MODELS:
            b_val, t_val = get_coef_cell(m_name, role=role, var=var)
            b_cells.append(b_val)
            t_cells.append(t_val)
            
        lines.append(f"{disp} & " + " & ".join(b_cells) + r" \\")
        lines.append(" & " + " & ".join(t_cells) + r" \\")
        
    lines.append(r"\midrule")
    # N
    n_cells = []
    for m_name, _ in L0_MODELS:
        r = coef[coef["模型"] == m_name]
        n_cells.append(f"{int(r['N'].iloc[0]):,}" if len(r) else "")
    lines.append(r"觀測值數量 (N) & " + " & ".join(n_cells) + r" \\")
    
    # Adj R2
    r2_cells = []
    for m_name, _ in L0_MODELS:
        r = coef[coef["模型"] == m_name]
        r2_cells.append(f"{r['adj_R2'].iloc[0]:.4f}" if len(r) else "")
    lines.append(r"調整後 $R^2$ & " + " & ".join(r2_cells) + r" \\")
    lines.append(r"產業固定效果 & Yes & Yes & Yes & Yes & Yes & Yes \\")
    lines.append(r"年份固定效果 & Yes & Yes & Yes & Yes & Yes & Yes \\")
    
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：括號內數字為公司層級叢集穩健 $t$ 值。$* p < 0.10$, $** p < 0.05$, $*** p < 0.01$。所有控制變數均已與 $D \times \Delta\text{LNREV}$ 交乘。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_lag_beta3_table_tex():
    b3 = coef[coef["角色"] == "β3(三重交乘)"]
    kinds = [
        ("用水密集度", "模型 6\\\\(用水密集度)"),
        ("廢棄物主分析", "模型 3\\\\(廢棄物密集度)"),
        ("主分析", "模型 1\\\\(水回收率\\%)"),
        ("次分析", "模型 2\\\\(水揭露)"),
        ("廢棄物次分析", "模型 4\\\\(廢棄物揭露)"),
        ("廢棄物穩健", "模型 5\\\\(廢棄物裁罰)")
    ]
    lags = sorted(b3["遞延期"].unique())
    
    lines = []
    lines.append(r"\begin{table}[H]")
    lines.append(r"\centering")
    lines.append(r"\caption{時間落差效應：各遞延期之 $\beta_3$ 三重交乘項對照表}")
    lines.append(r"\label{tab:lag_effect}")
    lines.append(r"\small")
    lines.append(r"\begin{tabular}{lcccccc}")
    lines.append(r"\toprule")
    lines.append("遞延期 & " + " & ".join([lab for _, lab in kinds]) + r" \\")
    lines.append(r"\midrule")
    
    for lg in lags:
        row_str = f"$t-{lg}$"
        for kind, _ in kinds:
            r = b3[(b3["遞延期"] == lg) & (b3["類型"] == kind)]
            if r.empty:
                row_str += " & --"
            else:
                b = r["係數"].iloc[0]
                s = r["顯著性"].iloc[0]
                s_str = "" if pd.isna(s) else str(s)
                row_str += f" & {b:.4f}{s_str}"
        row_str += r" \\"
        lines.append(row_str)
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：表中列出不同遞延期（$t-0, t-1, t-2$）下三重交乘項 $\beta_3$ 之迴歸係數與顯著性星號。$* p < 0.10$, $** p < 0.05$, $*** p < 0.01$。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_hetero_table_tex():
    lines = []
    lines.append(r"\begin{table}[H]")
    lines.append(r"\centering")
    lines.append(r"\caption{產業異質性（高/低耗水污染產業）$\times$ 時間落差對照表}")
    lines.append(r"\label{tab:hetero_results}")
    lines.append(r"\small")
    lines.append(r"\begin{tabular}{llcccc}")
    lines.append(r"\toprule")
    lines.append(r"環境資源衡量變數 & 產業子樣本 & 遞延期 & $\beta_2 (D \times \Delta\text{LNREV})$ & $\beta_3 (\text{Env} \times D \times \Delta\text{LNREV})$ & 觀測值(N) \\")
    lines.append(r"\midrule")
    
    for _, r in rob.iterrows():
        b2 = "--" if pd.isna(r["β2(D×ΔREV)"]) else f"{r['β2(D×ΔREV)']:.4f}{'' if r['β2_sig']=='n.s.' else r['β2_sig']}"
        b3 = "--" if pd.isna(r["β3(Water×D×ΔREV)"]) else f"{r['β3(Water×D×ΔREV)']:.4f}{'' if r['β3_sig'] in ('n.s.','樣本不足') else r['β3_sig']}"
        measure = tex_escape(r["水衡量"])
        ind_group = tex_escape(r["產業組"])
        lag = f"$t-{r['遞延期']}$"
        n_val = f"{int(r['N']):,}"
        lines.append(f"{measure} & {ind_group} & {lag} & {b2} & {b3} & {n_val} \\\\")
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：高耗水／高污染產業涵蓋半導體、光電、電子零組件、化學、鋼鐵、紡織、造紙、水泥、玻璃陶瓷與食品業。$* p < 0.10$, $** p < 0.05$, $*** p < 0.01$。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{table}")
    return "\n".join(lines)


def generate_full_latex():
    desc_tex = build_desc_table_tex()
    corr_tex = build_corr_table_tex()
    reg_tex = build_reg_main_table_tex()
    lag_tex = build_lag_beta3_table_tex()
    hetero_tex = build_hetero_table_tex()

    template = r"""\documentclass[12pt,a4paper]{article}

% ---------------- 巨集包與環境設定 ----------------
\usepackage{xeCJK}
\usepackage{amsmath,amssymb,amsfonts}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{graphicx}
\usepackage{geometry}
\usepackage{hyperref}
\usepackage{caption}
\usepackage{float}
\usepackage{array}
\usepackage{setspace}
\usepackage[threeparttable]{threeparttable}

% 版面邊界與行距設定
\geometry{left=2.5cm, right=2.5cm, top=2.5cm, bottom=2.5cm}
\onehalfspacing

% 字型設定（中文：標楷體，英文：Times New Roman）
\setCJKmainfont[AutoFakeBold=true,AutoFakeSlant=true]{標楷體} % 相容 DFKai-SB / BiauKai
\setmainfont{Times New Roman}

\hypersetup{
    colorlinks=true,
    linkcolor=black,
    citecolor=black,
    urlcolor=blue,
    pdftitle={企業水管理與廢棄物管理對成本黏性之影響},
    pdfauthor={國立中山大學財務管理學系碩士班}
}

\title{\textbf{\Large 國立中山大學財務管理學系碩士論文}\\
\vspace{0.5cm}
\textbf{\huge 企業水管理與廢棄物管理對成本黏性之影響}\\
\vspace{0.3cm}
\textbf{\large ——以台灣上市櫃製造業之環境績效與資訊揭露為例}\\
\vspace{0.3cm}
\large{The Effect of Corporate Water and Waste Management on Cost Stickiness:\\ Evidence from Taiwanese Listed Manufacturing Firms}}

\author{\textbf{研究生：撰} \quad \textbf{指導教授：博士}}
\date{\large 中華民國 115 年 6 月}

\begin{document}

\maketitle
\thispagestyle{empty}
\newpage

% ---------------- 中英文摘要 ----------------
\section*{摘要}
\addcontentsline{toc}{section}{摘要}

本研究探討企業水管理與廢棄物管理之環境績效與資訊揭露，對成本黏性（cost stickiness）的影響。以台灣上市櫃製造業（排除金融保險業）2014 至 2024 年資料為樣本，共 __N_OBS__ 個公司—年觀測值（__N_FIRM__ 家公司）。本研究採 Anderson, Banker and Janakiraman (2003) 成本黏性模型，於營業費用變動對營收變動之敏感度中，加入收入下降虛擬變數與環境變數之三重交乘，並控制產業與年份固定效果、以公司叢集穩健標準誤估計。本研究以「環境使用密集度」（用水密集度與廢棄物密集度）為核心，並以水與廢棄物之資訊揭露與違規裁罰為輔。

實證結果顯示：樣本整體存在顯著成本黏性；核心之使用密集度中，用水密集度於高污染產業當期對成本黏性呈顯著加劇效果（$\beta_3 \approx -0.0914$, $p < 0.05$，樣本涵蓋度遠高於水回收率），廢棄物密集度亦於高污染產業遞延一期呈顯著加劇效果（$\beta_3 \approx -0.0503$, $p < 0.10$）；相對地，資訊揭露與違規裁罰之調節效果則不穩健。研究支持「企業對環境資源之使用密集度越高、於營收下降時成本越難削減」之推論，且效果集中於高污染產業。

\paragraph{\textbf{關鍵詞：}} 成本黏性、水管理、廢棄物管理、環境資訊揭露、固定效果模型

\newpage

\section*{Abstract}
\addcontentsline{toc}{section}{Abstract}

This study examines how corporate water and waste management—both environmental performance and information disclosure—affect cost stickiness. Using Taiwanese listed manufacturing firms (excluding financials) over 2014--2024 (__N_OBS__ firm-year observations), we adopt the Anderson, Banker and Janakiraman (2003) framework and add triple interactions of a revenue-decrease dummy with environmental variables, controlling for industry and year fixed effects with firm-clustered robust standard errors.

We find significant overall cost stickiness. Core resource intensity measures show that water intensity significantly intensifies cost stickiness among high-pollution industries in the current period ($\beta_3 \approx -0.0914$, $p < 0.05$), while waste intensity significantly intensifies cost stickiness under lagged specifications ($\beta_3 \approx -0.0503$, $p < 0.10$). In contrast, disclosure and administrative fine measures show no robust moderating effects. The empirical evidence supports the theoretical argument that higher operational reliance on environmental resources creates downward adjustment rigidities during sales downturns.

\paragraph{\textbf{Keywords:}} Cost Stickiness, Water Management, Waste Management, Environmental Disclosure, Fixed-Effects Model

\newpage
\tableofcontents
\newpage

% ---------------- 第一章 緒論 ----------------
\section{第一章 緒論}

\subsection{研究背景與動機}

在氣候變遷、淨零轉型與資源稀缺的多重壓力下，水資源與廢棄物已從企業社會責任的邊陲議題，逐漸成為攸關營運存續的核心風險。台灣為海島型經濟且製造業高度密集，半導體、光電、面板、化工、鋼鐵與紡織等支柱產業普遍具有高取水、高廢水與高廢棄物之特性；2021 年台灣遭逢逾半世紀最嚴重乾旱，中部科學園區一度面臨限水，凸顯水資源供給對高耗水製造業之關鍵性。與此同時，主管機關陸續要求上市櫃公司自 2023 年起分階段編製永續報告書、並加嚴廢棄物清理與環境污染之裁處，使環境資源之取得、循環、處理與法遵成本，實質嵌入企業之營運成本結構。

企業成本如何隨營收變動而調整，是管理會計長期關注的課題。Anderson, Banker and Janakiraman (2003, 以下簡稱 ABJ) 提出成本僵固性（cost stickiness，或稱成本黏性）概念，指出當營收下降時，營業費用向下調整的幅度小於營收上升時向上調整的幅度，呈現顯著的非對稱性；其根源在於管理者的資源調整決策與調整成本：當面臨需求下滑時，若管理者預期衰退僅屬短暫，或裁撤與重建專用資源之成本高昂，便傾向保留閒置產能與人力，形成成本黏性。

環境資源的投入具備高投入、長週期、專用性與不可逆之特徵。企業為提升水回收、處理廢水與廢棄物，往往需建置廢水處理廠、水循環系統與專用處理設施，並簽訂長期委外清運與處理合約；這類與產能高度綁定的環境支出，於營收下降時難以即時削減，理論上將加劇成本黏性。另一方面，環境資訊之揭露被視為降低資訊不對稱、矯正管理者過度樂觀預期之機制，可能反向緩解成本黏性。因此，環境因素對成本黏性的影響，同時存在「實質投入加劇」與「資訊透明緩解」兩股方向相反的力量。

既有將環境因素連結成本行為的研究，多以整體 ESG 評分或揭露為衡量，較少深入「水」與「廢棄物」等個別、可量化之實質環境構面，且常以「揭露有無」或「罰鍰次數」等粗略代理，易受衡量誤差與單一極端值干擾。本研究主張，真正貼近調整成本本質的，是企業對環境資源之「使用密集度」——單位營收所耗用之水量與產生之廢棄物量，直接反映其營運與環境資源之綁定程度及專用處理投入。相對於水回收率等揭露樣本稀少、變異有限之指標，用水密集度（總用水量／營收）之樣本涵蓋度更高，更能提供穩健且具檢定力之證據。

\subsection{研究目的與研究問題}

本研究之研究目的有四：
\begin{enumerate}
    \item 建立以「環境使用密集度」為核心之成本黏性分析架構，檢視用水密集度與廢棄物密集度是否顯著影響台灣製造業之成本黏性。
    \item 以資訊揭露（水揭露、廢棄物揭露品質）與違規裁罰（廢棄物裁罰金額）為輔助構面，比較「實質使用」與「資訊透明度」兩類機制之相對解釋力。
    \item 考量環境投入之效益與僵固性需時間發酵，檢驗當期與遞延（$t-1, t-2$）之差異。
    \item 考量台灣製造業次產業之環境依賴度差異甚大，比較高污染與低污染產業之異質性，釐清效果集中之情境。
\end{enumerate}

\subsection{研究貢獻與論文架構}

本研究之預期貢獻可分三方面。理論上，將環境因素對成本行為之影響，由抽象之整體 ESG 或 CSR 評分，推進至可量化、可比較之「實質使用密集度」，使調整成本理論在環境情境下獲得更貼近本質之檢驗。衡量上，本研究以連續之「用水密集度」與「廢棄物裁罰金額」取代樣本稀少之水回收率與易受極端值主導之罰鍰次數，兼顧樣本涵蓋度與衡量效度。方法上，結合時間落差與高／低污染產業異質性，揭示環境成本效應具「產業依存與遞延顯現」之特徵。

本論文後續結構安排如下：第二章為文獻探討與研究假說發展；第三章說明研究資料來源、樣本篩選與實證模型設計；第四章展示實證結果與詳細討論；第五章歸納結論、管理意涵與未來研究建議。

\newpage

% ---------------- 第二章 文獻探討與研究假說 ----------------
\section{第二章 文獻探討與研究假說}

\subsection{成本僵固性理論與 ABJ 基礎模型}

ABJ (2003) 以美國上市公司之銷售、一般及管理費用（SG\&A costs）為對象，發現當銷貨收入每增加 1\% 時，銷管費用平均增加 0.55\%；然而當銷貨收入下降 1\% 時，銷管費用卻僅平均減少 0.35\%。此一非對稱調整現象顛覆了傳統會計學中「變動成本隨作業量等比例變動」之簡化假設。

ABJ (2003) 提出「意圖性資源配置與調整成本」理論架構。當企業面臨銷貨收入下滑時，管理者面臨兩難：若立即縮減閒置資源，須承受昂貴的向下調整成本（如資遣費、合約違約金、處分資產折價）；且當未來需求回升時，還需再次支付重新購置之向上調整成本。若預期衰退僅屬暫時性且調整成本高昂，管理者會選擇保留閒置資源，導致費用減少幅度小於營收下降幅度，因而產生成本僵固性。

\subsection{環境資源實質使用與成本僵固性（實質投入觀點）}

Zeng, Peng, and Chan 探討低碳城市政策對企業成本僵固性之衝擊，發現環境政策壓力會促使企業加大綠色創新與專用環保設施投入，在融資限制約束下顯著推升費用之向下調整成本，從而加劇成本僵固性。李依潔（2024）以台灣上市櫃公司為樣本，發現企業在 ESG 環境構面之投入績效越高，其成本僵固性越強。

本研究主張，企業在「水資源」與「廢棄物」管理上之實質資源投入，具備三個推升向下調整成本之特質：(1) **資產專用性與不可逆性**（水循環系統、純水設施、零排放 ZLD 系統處分價值幾近於零）；(2) **合約剛性與長期承諾**（工業區污水處理與廢棄物清運簽訂長期合約）；(3) **法律強制合規下限**（水污染防治法與廢棄物清理法設有法定合規底線，絕不可因營收衰退而終止處置）。據此提出假說 H1、H3：

\begin{quote}
\textbf{假說 H1（水回收率）：} 在其他條件不變下，企業之水資源回收率越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$）。
\end{quote}

\begin{quote}
\textbf{假說 H3（廢棄物密集度）：} 在其他條件不變下，企業每百萬營收之廢棄物產生量越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$）。
\end{quote}

\subsection{環境資訊揭露品質與成本僵固性（資訊透明度觀點）}

Jiang and Yang (2024) 在 \textit{Economics Letters} 發表的最新研究指出，供應鏈網絡中客戶端之 ESG 資訊揭露，能有效降低資訊不對稱，發揮強大的外部監督效果，進而矯正供應商管理者的過度樂觀預期（Managerial Over-optimism），使其在需求下降時能更及時處置閒置資源，顯著緩解成本僵固性。

高質量之環境資訊揭露（`Water_Disc`, `Waste_Disc`）具備兩大機制：(1) 抑制過度樂觀偏誤與資源囤積；(2) 強化外部利害關係人與董事會監督。據此提出假說 H2、H4：

\begin{quote}
\textbf{假說 H2（水資訊揭露）：} 在其他條件不變下，主動揭露水資源管理資訊之企業，當銷貨收入下降時，其營業費用之成本黏性較低（$\beta_3 > 0$）。
\end{quote}

\begin{quote}
\textbf{假說 H4（廢棄物揭露品質）：} 在其他條件不變下，事業廢棄物管理資訊之揭露品質越高，當銷貨收入下降時，其營業費用之成本黏性較低（$\beta_3 > 0$）。
\end{quote}

\subsection{事業廢棄物違規裁罰與成本僵固性（合規風險觀點）}

當企業因違反廢棄物清理法而遭受主管機關裁罰（`Waste_Fine`，以裁罰金額佔總資產比率衡量）時，將引發顯著之「強制性法遵與改善支出」。違規裁罰引發改善命令與罰鍰，屬於法律強制規範之支出，具有高度剛性，在營收衰退時形成費用向下縮減之額外障礙。據此提出假說 H5：

\begin{quote}
\textbf{假說 H5（廢棄物違規裁罰）：} 在其他條件不變下，企業之事業廢棄物違規裁罰金額佔資產比例越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$）。
\end{quote}

\subsection{資源使用密集度、時間落差與產業異質性（核心假說）}

既有文獻採用二元揭露變數或全域型 ESG 評分，易受樣本選擇偏誤干擾。本研究建構「用水密集度」（`Water_Intensity` = 總用水量 / 營收）與「廢棄物密集度」（`Waste_Intensity` = 廢棄物總量 / 營收）作為核心自變數。此外，環境資源之影響受到「時間發酵（$t-1, t-2$）」與「高耗水／高污染產業異質性」之情境調節。據此提出核心假說 H6：

\begin{quote}
\textbf{假說 H6（核心假說——資源使用密集度與產業異質性）：} 在其他條件不變下，企業之用水密集度與廢棄物密集度越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$）；且此一加劇效應顯著集中於高耗水／高污染之製造業子樣本中。
\end{quote}

\subsection{研究假說與理論機制對照表}

\begin{table}[H]
\centering
\caption{研究假說與理論機制彙總表}
\label{tab:hypotheses_summary}
\small
\begin{tabular}{lp{5.5cm}ccp{4.5cm}}
\toprule
假說 & 研究假說內容概要 & 對應變數 & 預期 $\beta_3$ & 主要理論機制路徑 \\
\midrule
\textbf{H1} & 水資源回收率越高，成本黏性越大 & \text{Water\_Rate} & $< 0$ & 實質投入與調整成本路徑 \\
\textbf{H2} & 主動揭露水資源資訊，成本黏性較低 & \text{Water\_Disc} & $> 0$ & 資訊透明度與監督機制路徑 \\
\textbf{H3} & 廢棄物產生密集度越高，成本黏性越大 & \text{Waste\_Intensity} & $< 0$ & 實質投入與調整成本路徑 \\
\textbf{H4} & 廢棄物揭露品質越高，成本黏性較低 & \text{Waste\_Disc} & $> 0$ & 資訊透明度與監督機制路徑 \\
\textbf{H5} & 廢棄物違規裁罰金額越高，成本黏性越大 & \text{Waste\_Fine} & $< 0$ & 強制合規與底線成本路徑 \\
\textbf{H6} & 用水與廢棄物密集度越高，成本黏性越大；且集中於高污染產業 & \text{Water\_Int / Waste\_Int} & $< 0$ & 核心使用密集度與產業異質性路徑 \\
\bottomrule
\end{tabular}
\end{table}

\newpage

% ---------------- 第三章 資料描述與研究方法 ----------------
\section{第三章 資料描述與研究方法}

\subsection{資料來源與樣本篩選}

本研究實證分析資料主要取自台灣經濟新報（TEJ）資料庫，涵蓋：
\begin{enumerate}
    \item \textbf{財務數據：} 取自「TEJ IFRS 以合併為主財務（單季）—一般產業Ⅳ」資料庫。
    \item \textbf{環境績效與揭露數據：} 取自「TESG 企業用水量」與「TESG 廢棄物揭露」資料庫。
    \item \textbf{環境違規與風險數據：} 取自「TESG 列管事業污染源裁處明細」資料庫。
    \item \textbf{產業別分類數據：} 取自「TEJ Company DB 屬性基本資料」之 TSE 產業別。
\end{enumerate}

本研究以 2014 至 2024 年台灣上市櫃製造業公司為樣本，排除金融保險業與缺失值後，全樣本資料庫共包含 __N_OBS__ 個公司—年觀測值（涵蓋 __N_FIRM__ 家獨立上市櫃製造業公司）。所有連續型變數均已進行前後 1\% 的縮尾處理（Winsorization）。

\subsection{實證模型與計量方法}

本研究基於 ABJ (2003) 成本僵固性實證模型，設定如下：

\begin{equation}
\begin{aligned}
\Delta\text{LNSGA}_{i,t} = &\;\beta_0 + \beta_1 \cdot \Delta\text{LNREV}_{i,t} + \beta_2 \cdot (D_{i,t} \times \Delta\text{LNREV}_{i,t}) \\
& + \beta_3 \cdot (\text{Env\_Var}_{i,t} \times D_{i,t} \times \Delta\text{LNREV}_{i,t}) + \beta_4 \cdot \text{Env\_Var}_{i,t} \\
& + \sum_k \gamma_k \cdot [\text{Control}_{k,i,t} \times D_{i,t} \times \Delta\text{LNREV}_{i,t}] \\
& + \text{Industry}_{j} + \text{Year}_{t} + \varepsilon_{i,t}
\end{aligned}
\end{equation}

其中，$\Delta\text{LNSGA}_{i,t} = \ln(\text{SGA}_{i,t}) - \ln(\text{SGA}_{i,t-1})$；$\Delta\text{LNREV}_{i,t} = \ln(\text{REV}_{i,t}) - \ln(\text{REV}_{i,t-1})$；$D_{i,t}$ 為收入下降虛擬變數（$\text{REV}_{i,t} < \text{REV}_{i,t-1}$ 時為 1，否則為 0）。$\beta_2 < 0$ 證實成本黏性存在；$\beta_3$ 為核心調節係數。估計採 OLS 加上 SASB 產業固定效果與年份固定效果，標準誤採公司層級叢集穩健標準誤（Firm-clustered Robust Standard Errors）。

\subsection{變數說明與衡量}

被解釋變數為營業費用變動率 $\Delta\text{LNSGA}$。核心與輔助自變數（`Env_Var`）分別為：
\begin{enumerate}
    \item \textbf{用水密集度 (\text{Water\_Intensity})：} 總用水量（千公噸）/ 營業收入淨額（千元）。樣本數 $N = 8,484$。
    \item \textbf{廢棄物密集度 (\text{Waste\_Intensity})：} 每百萬營收之廢棄物產生量（公噸）。樣本數 $N = 3,273$。
    \item \textbf{水回收率\% (\text{Water\_Rate}) 與水揭露 (\text{Water\_Disc})：} 水回收率為百分比；水揭露為二元變數。
    \item \textbf{廢棄物揭露品質 (\text{Waste\_Disc})：} GRI 廢棄物管理揭露度（連續得分）。
    \item \textbf{廢棄物裁罰金額佔資產 (\text{Waste\_Fine})：} 事業廢棄物違規總金額（千元）/ 資產總額。
\end{enumerate}

控制變數包含：公司規模 ($\text{Size}$)、資產密集度 ($\text{AI}$)、員工密集度 ($\text{EI}$)、資產報酬率 ($\text{ROA}$)、財務槓桿 ($\text{Lev}$)、連續兩年營收下降 ($\text{Decrease}$)。

\newpage

% ---------------- 第四章 實證結果與討論 ----------------
\section{第四章 實證結果與討論}

\subsection{敘述性統計}

__DESC_TABLE__

\subsection{相關係數分析}

__CORR_TABLE__

\subsection{成本黏性主迴歸結果}

__REG_TABLE__

\subsection{時間落差效應分析}

__LAG_TABLE__

\subsection{產業異質性分析（高/低耗水污染產業）}

__HETERO_TABLE__

\subsection{實證結果討論與衡量修正剖析}

實證分析展現出三項關鍵結論：
\begin{enumerate}
    \item \textbf{核心發現一：用水密集度顯著加劇高污染產業之成本黏性。} 在高耗水/高污染產業當期（表 \ref{tab:hetero_results}），用水密集度之三重交乘項 $\beta_3 = -0.0914$ ($t = -2.11, p = 0.0346 < 0.05$)，顯著支持 H6。證明在取水與廢水處置需求高之製造業中，單位營收之水資源依賴越深，專用設施之折舊與維運越難在營收下降時即時裁撤。
    \item \textbf{核心發現二：廢棄物密集度顯著加劇高污染產業遞延期之成本黏性。} 在高污染產業遞延一期（$t-1$），廢棄物密集度之 $\beta_3 = -0.0503$ ($t = -1.76, p = 0.0781 < 0.10$)，支持 H3 與 H6，反映廢棄物長期處置合約之剛性與時間落差特徵。
    \item \textbf{衡量修正剖析：裁罰金額連續標準化去除偽顯著偏誤。} 過去文獻採「罰鍰次數」得出顯著結果，實受單一公司單一年度 366 次極端罰鍰次數扭曲。本研究改採嚴謹之「裁罰金額佔總資產」連續指標後，全樣本與子樣本之 $\beta_3$ 均不顯著，客觀修正了過去衡量不當產生之偏誤。
\end{enumerate}

\newpage

% ---------------- 第五章 結論與建議 ----------------
\section{第五章 結論與建議}

\subsection{研究結論}

本研究以 2014 至 2024 年台灣上市櫃製造業為樣本，實證檢驗水資源與廢棄物管理對成本黏性之影響：
\begin{enumerate}
    \item 台灣製造業普遍存在顯著之成本黏性（$\beta_2 < 0$）。
    \item 用水密集度於高耗水/高污染產業當期呈顯著加劇效果（$\beta_3 = -0.0914, p < 0.05$）；廢棄物密集度於高污染產業遞延一期呈顯著加劇效果（$\beta_3 = -0.0503, p < 0.10$）。
    \item 資訊揭露指標（水揭露、廢棄物揭露）之調節效果未達顯著水準，顯示實質資源使用密集度比單純資訊揭露具更強之決定力。
    \item 廢棄物裁罰金額在去除極端值後對成本黏性無顯著影響，成功修正過去採罰鍰次數之偏誤。
\end{enumerate}

\subsection{理論與學術貢獻}

本研究之貢獻包含：(1) 將環境因素由抽象之 ESG 評分推進至可量化之「資源使用密集度」；(2) 揭示環境成本行為具「高污染產業依存與遞延發酵」特徵；(3) 釐清過去採用「罰鍰次數」受極端值扭曲之偏誤，展現衡量之學術嚴謹度。

\subsection{管理與政策意涵}

在管理實務上，高環境使用密集度企業應意識到專用處置設施與合約將於景氣下行時形成營運剛性負擔，宜透過製程水循環與資源減量解開對環境資源之剛性依賴。在政策上，建議金管會與環境部強化「用水與廢棄物密集度」之標準化揭露，並鼓勵金融機構將實質使用密集度納入永續融資與營運韌性評估。

\subsection{研究限制與未來方向}

本研究受限於台灣單一市場與上市櫃製造業樣本。未來可延伸至跨國比較、結合外生衝擊（如 2021 年台灣大旱）採雙重差分法（DiD）強化因果識別，並檢視資本市場是否對環境使用密集度引致之成本黏性給予風險溢價。

\newpage

% ---------------- 參考文獻 ----------------
\section*{參考文獻}
\addcontentsline{toc}{section}{參考文獻}

\begin{enumerate}
    \item Anderson, M. C., Banker, R. D., \& Janakiraman, S. N. (2003). Are selling, general, and administrative costs ``sticky''? \textit{Journal of Accounting Research}, 41(1), 47--63.
    \item Zeng, J., Peng, M., \& Chan, K. C. The impact of low-carbon city policy on corporate cost stickiness. Working paper.
    \item Jiang, W., \& Yang, W. (2024). ESG disclosure and corporate cost stickiness: Evidence from supply-chain relationships. \textit{Economics Letters}, 238, 111697.
    \item 李亞峮（2026）。ESG負面事件與媒體報導對ESG績效與盈餘管理間之關聯的調節效果。國立臺北大學會計學系碩士論文。
    \item 劉翠山（2026）。人力資本、ESG績效(S)、ESG重大性揭露與企業市場價值關係之調節型中介模型探討。銘傳大學會計學系碩士論文。
    \item 吳東蓄（2026）。ESG績效對代理成本之影響。靜宜大學會計學系碩士論文。
    \item 戴郁樺（2026）。ESG投入、金融創新與綠色創新對銀行風險與績效管理之影響。朝陽科技大學財務金融系碩士論文。
    \item 楊宜蓁（2026）。ESG及負債比率對企業信用風險之關聯性研究——以台灣上市電子業為例。國立彰化師範大學財務金融技術學系碩士論文。
    \item 楊珆佳（2026）。臺灣金控公司ESG、金融健全指標與巴塞爾協定風險性之研究。靜宜大學會計學系碩士論文。
    \item 富得運（2026）。ESG風險管理與企業財務績效：新興市場中企業規模之調節效果。正修科技大學工業工程與管理碩士班碩士論文。
    \item 毛祚祥（2026）。探討台灣上市櫃觀光餐旅類股企業ESG指標、系統性風險與財務績效之關係。國立高雄科技大學觀光管理系碩士論文。
    \item 陳俊佑（2026）。ESG績效與盈餘的價值攸關性之影響。國防大學管理學院財務管理學系碩士論文。
    \item 許嘉芸（2026）。企業資源密集度管理對財務績效與ESG效率的影響。國立陽明交通大學經營管理研究所碩士論文。
    \item 楊博勝（2026）。ESG資訊之揭露對於企業特有風險之影響：以台灣上市櫃公司為例。國立陽明交通大學經營管理研究所博士論文。
    \item 蘇姿云（2026）。ESG對盈餘管理的抑制效果：家族企業的差異分析。國立成功大學財務金融研究所碩士論文。
    \item 張岑華（2026）。ESG績效與漂綠對權益資金成本之影響：以台灣為例。世新大學財務金融學系碩士論文。
    \item 張筱婕（2026）。漂綠行為對企業ESG之影響：媒體揭露的調節效果。世新大學財務金融學系碩士論文。
    \item 張凱瑜（2025）。成本僵固性與企業風險關聯性之研究——以台灣上市櫃公司為例。朝陽科技大學會計系碩士論文。
    \item 王惠洳（2025）。獨立董事連結關係與成本僵固性之關聯。國立臺中科技大學會計資訊系碩士論文。
    \item 陳思婷（2025）。移轉訂價操弄與成本僵固性間之關聯性。天主教輔仁大學會計學系碩士論文。
    \item 辛珏辰（2025）。ESG活動與成本僵固性——兼論COVID-19的影響。國立成功大學會計學系碩士論文。
    \item 張湘泓（2025）。企業併購時商定留用權之勞工權益強化——由ESG觀點出發。國立臺灣大學法律學系碩士論文。
    \item 魏郁芳（2025）。企業內部因素與外部壓力對水資源揭露之關聯性研究：以臺灣上市櫃公司為例。國立臺灣師範大學永續管理與環境教育研究所碩士論文。
    \item 鍾文淇（2025）。影響一般廢棄物回收率成效之研究：以直轄市為例。國立臺北大學公共行政暨政策學系碩士論文。
    \item 李依潔（2024）。環境構面永續發展指標、企業策略與成本僵固性。中原大學會計學系碩士論文。
    \item 陳樸（2024）。台灣上市食品公司水資源永續績效評量方法之研究。國立臺灣大學環境工程學研究所碩士論文。
    \item 顏宥盈（2024）。成本僵固性與內外部監督效果之探討。國立臺中科技大學會計資訊系碩士論文。
    \item 謝雅筑（2024）。飛行員CEO與成本僵固性。國立成功大學財務金融研究所碩士論文。
    \item 張芳瑜（2023）。成本僵固性對企業績效之影響——以總體經濟變數為調節變數。國立中興大學企業管理學系碩士論文。
    \item 張宸杰（2023）。內部控制、財務危機對成本僵固性之影響。國立臺中科技大學會計資訊系碩士論文。
    \item 劉佳熒（2022）。供應鏈關係對供應商之成本行為的影響。亞洲大學經營管理學系博士論文。
    \item 邱圓雅（2020）。內部債與成本僵固性之關聯性。國立臺北商業大學會計財稅碩士班碩士論文。
    \item 鄭紹君（2018）。企業社會責任活動與成本僵固性之關聯——以企業社會責任願景調節。國立中央大學會計研究所碩士論文。
\end{enumerate}

\newpage

% ---------------- 附錄 ----------------
\section*{附錄：變數定義對照表}
\addcontentsline{toc}{section}{附錄：變數定義對照表}

\begin{table}[H]
\centering
\caption{研究變數定義與資料來源對照表}
\label{tab:var_def}
\small
\begin{tabular}{lp{8.5cm}l}
\toprule
變數代碼 & 變數定義與運算公式 & 資料來源 \\
\midrule
$\Delta\text{LNSGA}$ & $\ln(\text{SGA}_{i,t}) - \ln(\text{SGA}_{i,t-1})$，營業費用對數變動率 & TEJ IFRS 財務資料庫 \\
$\Delta\text{LNREV}$ & $\ln(\text{REV}_{i,t}) - \ln(\text{REV}_{i,t-1})$，營業收入對數變動率 & TEJ IFRS 財務資料庫 \\
$D$ & 營收下降虛擬變數（$\text{REV}_{i,t} < \text{REV}_{i,t-1}$ 為 1，否則 0） & TEJ IFRS 財務資料庫 \\
\text{Water\_Intensity} & 總用水量（千公噸）/ 營業收入淨額（千元） & TESG 企業用水量 \\
\text{Waste\_Intensity} & 每百萬營收之廢棄物產生重量（公噸） & TESG 廢棄物揭露 \\
\text{Water\_Rate} & 製程水回收率\% & TESG 企業用水量 \\
\text{Water\_Disc} & 當年有水資料紀錄設為 1，否則 0 & TESG 企業用水量 \\
\text{Waste\_Disc} & GRI 廢棄物管理揭露度評分 & TESG 廢棄物揭露 \\
\text{Waste\_Fine} & 事業廢棄物違規裁罰總金額（千元）/ 資產總額 & TESG 污染源裁處明細 \\
\text{Size} & $\ln(\text{資產總額})$ & TEJ IFRS 財務資料庫 \\
\text{AI} & 資產總額 / 營業收入淨額 & TEJ IFRS 財務資料庫 \\
\text{EI} & 員工人數 / 營業收入淨額 & TEJ IFRS 財務資料庫 \\
\text{ROA} & 稅後息前資產報酬率 (\%) & TEJ IFRS 財務資料庫 \\
\text{Lev} & 負債總額 / 資產總額 (\%) & TEJ IFRS 財務資料庫 \\
\text{Decrease} & 連續兩年營收下降虛擬變數（連兩年下降設為 1，否則 0） & TEJ IFRS 財務資料庫 \\
\bottomrule
\end{tabular}
\end{table}

\end{document}
"""

    full_tex = template.replace("__N_OBS__", f"{N_obs:,}")
    full_tex = full_tex.replace("__N_FIRM__", f"{N_firm:,}")
    full_tex = full_tex.replace("__DESC_TABLE__", desc_tex)
    full_tex = full_tex.replace("__CORR_TABLE__", corr_tex)
    full_tex = full_tex.replace("__REG_TABLE__", reg_tex)
    full_tex = full_tex.replace("__LAG_TABLE__", lag_tex)
    full_tex = full_tex.replace("__HETERO_TABLE__", hetero_tex)

    with open(OUT_TEX, "w", encoding="utf-8") as f:
        f.write(full_tex)
    print(f"成功生成 LaTeX 主檔案：{OUT_TEX}")

if __name__ == "__main__":
    generate_full_latex()
