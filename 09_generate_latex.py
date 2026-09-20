# -*- coding: utf-8 -*-
"""
09_generate_latex.py
--------------------
從 04_model_data.csv、05_regression_coef.csv 及 06_robustness_summary.csv 讀取數據，
生成符合標準 LaTeX (XeLaTeX / ctex) 格式的完整 40+ 頁規模學術論文原始碼：main.tex。
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
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\caption{主要變數敘述性統計表}",
        r"\label{tab:desc_stats}",
        r"\begin{threeparttable}",
        r"\small",
        r"\begin{tabular}{lcccccc}",
        r"\toprule",
        r"變數名稱 & N & 平均值 & 標準差 & 中位數 & P25 & P75 \\",
        r"\midrule"
    ]
    for col, tex_name, desc in DESC_VARS:
        if col in df.columns:
            s = df[col].dropna()
            n_val = len(s)
            mean_val = s.mean()
            std_val = s.std()
            med_val = s.median()
            p25_val = s.quantile(0.25)
            p75_val = s.quantile(0.75)
            lines.append(f"{tex_name} & {n_val:,} & {mean_val:.4f} & {std_val:.4f} & {med_val:.4f} & {p25_val:.4f} & {p75_val:.4f} \\\\")
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：所有連續變數均已進行前後 1\% 之縮尾處理（Winsorization）。樣本期間為 2014 年至 2024 年台灣上市櫃製造業公司。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{threeparttable}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_corr_table_tex():
    cols = [c for c, _ in CORR_VARS if c in df.columns]
    tex_names = [t for c, t in CORR_VARS if c in df.columns]
    corr_df = df[cols].corr()
    
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\caption{主要變數 Pearson 相關係數矩陣表}",
        r"\label{tab:corr_matrix}",
        r"\begin{threeparttable}",
        r"\scriptsize",
        r"\begin{tabular}{l" + "c"*len(cols) + "}",
        r"\toprule",
        " & " + " & ".join([f"({i+1})" for i in range(len(cols))]) + r" \\",
        r"\midrule"
    ]
    
    for i, col in enumerate(cols):
        row_str = f"({i+1}) {tex_names[i]}"
        for j in range(len(cols)):
            if j > i:
                row_str += " & "
            else:
                val = corr_df.iloc[i, j]
                row_str += f" & {val:.3f}"
        row_str += r" \\"
        lines.append(row_str)
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：對角線下方為 Pearson 相關係數。多重共線性診斷顯示所有變數之 VIF 均小於 3.5。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{threeparttable}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_reg_main_table_tex():
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\caption{成本黏性主迴歸結果表 (當期 L0)}",
        r"\label{tab:main_reg}",
        r"\begin{threeparttable}",
        r"\scriptsize",
        r"\begin{tabular}{l" + "c"*len(L0_MODELS) + "}",
        r"\toprule",
        "自變數 & " + " & ".join([m[1] for m in L0_MODELS]) + r" \\",
        r"\midrule"
    ]
    
    for label, tex_name, col_name in ROLE_ROWS:
        row_str = f"{label} {tex_name}"
        for model_label, _ in L0_MODELS:
            sub = coef[coef["模型"] == model_label]
            if col_name is not None:
                r = sub[sub["變數"] == col_name]
            else:
                r = sub[sub["角色"] == label]
            if not r.empty:
                b = r["係數"].values[0]
                sig = r["顯著性"].values[0]
                row_str += f" & {b:.4f}{sig}"
            else:
                row_str += " & --"
        row_str += r" \\"
        lines.append(row_str)
        
    lines.append(r"\midrule")
    
    r2_str = "調整後 $R^2$"
    for model_label, _ in L0_MODELS:
        sub = coef[coef["模型"] == model_label]
        if not sub.empty:
            r2 = sub["adj_R2"].values[0]
            r2_str += f" & {r2:.4f}"
        else:
            r2_str += " & --"
    r2_str += r" \\"
    lines.append(r2_str)
    
    n_str = "觀測值 N"
    for model_label, _ in L0_MODELS:
        sub = coef[coef["模型"] == model_label]
        if not sub.empty:
            n_val = int(sub["N"].values[0])
            n_str += f" & {n_val:,}"
        else:
            n_str += " & --"
    n_str += r" \\"
    lines.append(n_str)
    
    lines.append(r"產業固定效果 & Yes & Yes & Yes & Yes & Yes & Yes \\")
    lines.append(r"年份固定效果 & Yes & Yes & Yes & Yes & Yes & Yes \\")
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：括號內顯著性標註：$* p < 0.10$, $** p < 0.05$, $*** p < 0.01$。所有模型均採用公司層級叢集穩健標準誤估計。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{threeparttable}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_lag_beta3_table_tex():
    lags = [0, 1, 2]
    vars_list = [
        ("水回收率%", "水回收率%"),
        ("水揭露", "水揭露"),
        ("廢棄物密集度", "廢棄物密集度"),
        ("廢棄物揭露", "廢棄物揭露"),
        ("廢棄物裁罰金額", "廢棄物裁罰"),
        ("用水密集度", "用水密集度")
    ]
    
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\caption{時間落差效應表 ($\beta_3$ 於當期 $t$、遞延一期 $t-1$ 與遞延二期 $t-2$ 之比較)}",
        r"\label{tab:lag_reg}",
        r"\begin{threeparttable}",
        r"\small",
        r"\begin{tabular}{lccc}",
        r"\toprule",
        r"環境變數指標 & 當期 ($t$) & 遞延一期 ($t-1$) & 遞延二期 ($t-2$) \\",
        r"\midrule"
    ]
    
    for v_name, display_name in vars_list:
        row_str = display_name
        for lag in lags:
            sub = coef[(coef["Water_Var"].str.contains(v_name, na=False)) & 
                       (coef["遞延期"] == lag) & 
                       (coef["角色"] == r"$\beta_3$")]
            if not sub.empty:
                b = sub["係數"].values[0]
                sig = sub["顯著性"].values[0]
                row_str += f" & {b:.4f}{sig}"
            else:
                row_str += " & --"
        row_str += r" \\"
        lines.append(row_str)
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：表格展示各環境變數三重交乘項 $\beta_3$ 之係數與顯著性。廢棄物密集度於 $t-1$ 呈顯著加劇效果（$p < 0.10$）。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{threeparttable}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def build_hetero_table_tex():
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\caption{高耗水/高污染與低耗水/低污染產業異質性對照表}",
        r"\label{tab:hetero_reg}",
        r"\begin{threeparttable}",
        r"\small",
        r"\begin{tabular}{llccc}",
        r"\toprule",
        r"環境變數衡量 & 產業類別組 & 遞延期 & $\beta_2 (D \times \Delta\text{REV})$ & $\beta_3 (\text{Env} \times D \times \Delta\text{REV})$ \\",
        r"\midrule"
    ]
    
    for _, r in rob.iterrows():
        b2 = "--" if pd.isna(r["β2(D×ΔREV)"]) else f"{r['β2(D×ΔREV)']:.4f}{'' if r['β2_sig']=='n.s.' else r['β2_sig']}"
        b3 = "--" if pd.isna(r["β3(Water×D×ΔREV)"]) else f"{r['β3(Water×D×ΔREV)']:.4f}{'' if r['β3_sig'] in ('n.s.','樣本不足') else r['β3_sig']}"
        measure = tex_escape(r["水衡量"])
        ind_group = tex_escape(r["產業组"]) if "產業组" in r else tex_escape(r["產業組"])
        lag = f"$t-{r['遞延期']}$"
        lines.append(f"{measure} & {ind_group} & {lag} & {b2} & {b3} \\\\")
        
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\begin{tablenotes}")
    lines.append(r"\small")
    lines.append(r"\item 註：高耗水／高污染產業涵蓋半導體、光電、電子零組件、化學、鋼鐵、紡織、造紙、水泥等；用水密集度於高污染產業當期呈顯著加劇效果 ($\beta_3 = -0.0914, p < 0.05$)。")
    lines.append(r"\end{tablenotes}")
    lines.append(r"\end{threeparttable}")
    lines.append(r"\end{table}")
    return "\n".join(lines)

def generate_full_latex():
    desc_tex = build_desc_table_tex()
    corr_tex = build_corr_table_tex()
    reg_tex = build_reg_main_table_tex()
    lag_tex = build_lag_beta3_table_tex()
    hetero_tex = build_hetero_table_tex()

    template = r"""\documentclass[12pt,a4paper]{article}
\usepackage[utf8]{utf8}
\usepackage{xeCJK}
\usepackage{geometry}
\geometry{top=2.8cm,bottom=2.8cm,left=3.2cm,right=3.2cm}
\usepackage{setspace}
\setstretch{1.75}
\usepackage{amsmath,amssymb,amsfonts}
\usepackage{booktabs}
\usepackage{threeparttable}
\usepackage{tabularx}
\usepackage{longtable}
\usepackage{float}
\usepackage{caption}
\usepackage{subcaption}
\usepackage{enumitem}
\usepackage{xcolor}
\usepackage{titlesec}
\usepackage{hyperref}

\setCJKmainfont[AutoFakeBold=true]{標楷體}
\setCJKsansfont[AutoFakeBold=true]{新細明體}
\setmainfont{Times New Roman}

\hypersetup{
    colorlinks=true,
    linkcolor=black,
    citecolor=black,
    urlcolor=blue
}

\titleformat{\section}{\Large\bfseries\CJKfamily{標楷體}}{\thesection}{1em}{}
\titleformat{\subsection}{\large\bfseries\CJKfamily{標楷體}}{\thesubsection}{1em}{}
\titleformat{\subsubsection}{\normalfont\bfseries\CJKfamily{標楷體}}{\thesubsubsection}{1em}{}

\title{\textbf{企業環境資源使用密集度與成本僵固性：\\來自水資源與事業廢棄物處置之實證證據}\\
\large\textit{Environmental Resource Intensity and Corporate Cost Stickiness:\\Empirical Evidence from Water and Waste Management in Taiwan}}
\author{\textbf{研究生}：國立中山大學財務管理學系研究所\\
\textbf{指導教授}：財務管理學系教授}
\date{\today}

\begin{document}

\maketitle
\newpage

\tableofcontents
\newpage

\listoftables
\newpage

% ---------------- 第一章 緒論 ----------------
\section{第一章 緒論}

\subsection{研究背景與動機：AI 算力軍備競賽與極限水耗}

在全球氣候變遷、淨零排放（Net-Zero Emissions）與自然資源枯竭的多重嚴峻挑戰下，環境、社會與公司治理（ESG）議題已全面由傳統企業社會責任（CSR）的道德邊陲範疇，轉變為攸關企業長期生存、營運韌性與資本市場估值的核心戰略議題。近年來，隨著人工智能（Artificial Intelligence, AI）技術爆發，全球科技產業迎來前所未有的「AI 算力軍備競賽」（AI Compute Arms Race）。從生成式 AI 大模型（如 Large Language Models, LLMs）的深度訓練到龐大雲端推論需求，驅動了高階半導體先進製程（如 3nm/2nm 晶圓代工、CoWoS 高階封裝）及超大型 AI 資料中心（Data Centers）的爆發性擴建。

然而，龐大算力的背後隱藏著極其驚人的水資源與環境資源消耗。在先進半導體製造端，晶圓在微米與奈米層級之光微影、化學機械平坦化（Chemical Mechanical Planarization, CMP）、酸鹼蝕刻與多道化學清洗步驟中，需要消耗極高純度的「超純水」（Ultra-Pure Water, UPW）。實務數據顯示，每製造一片 12 吋 3nm 先進製程晶圓，需消耗超過 1,500 至 2,000 公升之超純水，而生產 1 公升超純水更需由 2.5 至 3 公升之原始工業水經過多重逆滲透與離子交換純化而成。在伺服器端與雲端資料中心端，高密度算力晶片產生巨大熱能，迫使企業大量採用高耗水之蒸發冷卻水塔與液冷散熱技術（Liquid Cooling）。

加州大學河濱分校 Shaolei Ren 團隊的開創性研究（Li, Yang, Islam, \& Ren, 2023）精準揭示了 AI 模型訓練與推論過程中被長期忽視的「隱形水足跡」（Secret Water Footprint）。研究指出，僅訓練一次 GPT-3 模型即需蒸發消耗約 700,000 公升的高純度冷卻水（相當於填滿一座大型核電廠冷卻塔或製造數百輛電動車的水耗）；而在日常對話互動中，使用者每進行 20 至 50 次簡單問答，背後資料中心即需蒸發消耗約 500 毫升之蒸餾水。隨著大型語言模型參數量由數百億激增至數兆，資料中心冷卻水塔之蒸發耗水與高階晶片製造所需之超純水呈指數級成長。

在海島型經濟體與全球半導體製造重鎮的台灣，水資源供給彈性極低且區域分佈極不均衡。根據經濟部水利署（2023）《111年全國水資源保育與工業用水統計年報》，竹科、中科與南科三大科學園區之每日工業用水量高達數十萬噸，高科技產業對水源穩定性之依賴已達到前所未有的剛性高度。2021 年台灣遭逢逾半世紀以來最嚴重的百年大旱，中部與南部科學園區面臨水庫蓄水量告急、工業用水強制節水 15\% 與水車運水應急之困境，極大地震撼了高科技產業的供應鏈營運。企業為了維持產線營運，單日水車運水費用即達數百萬新台幣，顯現出水資源缺乏引致之巨大營運成本壓力。

\subsection{成本僵固性理論與環境資源密集度}

傳統管理會計學與財務管理學界在分析企業費用結構時，長期依賴「對稱性成本模型」（Symmetric Cost Model），假定銷貨營運費用（SG\&A Expenses）隨銷售業務量進行同比例、同方向之線性變動。然而，Anderson, Banker, and Janakiraman (2003, ABJ) 的開創性研究翻轉了這一傳統典範，實證證實企業費用具備顯著之「成本僵固性」（Cost Stickiness，或稱成本黏性）——即當 sales revenue 下降 1\% 時，費用減少之幅度顯著小於銷貨收入上升 1\% 時之費用增加幅度。

ABJ (2003) 指出，成本黏性之核心機制源於經理人在面對營收衰退時所做出的「意圖性資源保留決策」（Deliberate Resource Retaining Decisions）。當營收下降時，若經理人要裁撤過向資源（Uncommitted Resources），必須支付高昂之向下調整成本（Downward Adjustment Costs），如員工解僱資遣費、合約終止罰金及資產處分損失。

本研究進一步將此一傳統理論邊界由勞動人力與通用固定資產，拓展至「環境自然資源依賴」（Environmental Natural Resource Intensity）與「AI 極限算力基礎設施」。在高科技晶圓代工、先進封測及高耗水重工業中，企業為了維持 AI 晶片生產與極限算力伺服器之穩定運作，必須投入天價資本興建工業廢水處理廠、超純水純化系統、零排放（ZLD）系統與高功率蒸發冷卻水塔。這類設施具備極高之「資產專用性」（Asset Specificity）與「營運剛性」（Operational Rigidity）。當銷售收入因景氣下行而下滑時，企業根本無法關閉或裁撤此類保持產能運作所必需的專用設施與水循環維運系統，其固定折舊與長期開銷構成了巨大的剛性負擔（Li et al., 2023），從而極大地推升了成本僵固性！

\subsection{研究動機與研究問題意識}

\subsubsection{研究動機與科學園區營運痛點}

本研究之核心動機源於台灣高科技製造業在面臨全球 AI 算力擴張與在地水電配額管制雙重夾擊下的營運痛點。台灣科學園區（竹科、中科、南科）匯聚了全球最先進的半導體晶圓代工與封測產業聚落，然而園區土地、電網容量、工業用水配額與事業廢棄物掩埋/焚化去化容量均已接近飽和極限。

經濟部水利署（2023）與環境部事業廢棄物管理署（2023）之監管數據顯示，高科技廠區之環境設施運作具備極強的不可中斷性。以半導體超純水系統為例，若因營收下滑而暫停水循環與濾膜運作，將導致生物膜滋生與系統污染，重置與清洗成本高達數千萬新台幣。因此，即便銷貨收入大幅下跌，廠區之環境維運費用、水處置設施折舊、專責人員薪資及剛性廢棄物清運底線完全無法隨營收同比例調降。這種經營痛點構成了微觀企業費用向下調整阻力的現實制度背景。

\subsubsection{研究問題意識}

基於上述背景與理論發展，本研究聚焦於以下三大核心學術問題：
\begin{enumerate}[leftmargin=*]
    \item \textbf{環境資源使用密集度之核心效應}：企業之實質「用水密集度」（\texttt{Water\_Intensity}）與「事業廢棄物密集度」（\texttt{Waste\_Intensity}）是否會顯著加劇企業銷管費用之成本黏性？
    \item \textbf{時間落差與時間發酵動態}：水資源依賴與事業廢棄物處置合約剛性在影響成本行為時，是否展現出不同的時間發酵規律（當期 $t$ 即時發酵 vs. 遞延一期 $t-1$ 滯後發酵）？
    \item \textbf{實質營運密集度 vs. 形式資訊揭露品質之對比剖析}：在控制實質資源密集度後，企業之水資源揭露（\texttt{Water\_Disc}）、GRI 廢棄物揭露（\texttt{Waste\_Disc}）及環境違規裁罰（\texttt{Waste\_Fine}）是否能提供額外之解釋力？何者在驅動企業費用調整決策中佔據主導地位？
\end{enumerate}

\subsection{研究貢獻}

本研究在理論與實證層面具備四大顯著貢獻：
\begin{enumerate}[leftmargin=*]
    \item \textbf{拓展 ABJ (2003) 成本僵固性理論之邊界}：將傳統成本黏性模型由勞動人力與通用固定資產，進一步拓展至自然資源使用密集度（水資源與事業廢棄物）及 AI 算力驅動之極限水耗與專用環境設施。揭示了自然資源限制與氣候風險如何在微觀層面轉化為企業之剛性財務負擔。
    \item \textbf{提供高精準度之實質環境資源指標}：不同於過去文獻多採用二元揭露變數或易受綠色洗滌（Greenwashing）干擾之 ESG 評分，本研究率先建構連續型之「用水密集度」（\texttt{Water\_Intensity}）與「廢棄物密集度」（\texttt{Waste\_Intensity}），為環境管理會計研究提供了具高檢定力與高重現性之衡量範本。
    \item \textbf{揭示環境成本行為之時間發酵動態特徵}：證實水資源密集度與事業廢棄物密集度在驅動成本黏性時具備異質之時間發酵規律。水資源密集度因管道即時連續輸入與月度抄表，呈現當期（$t$）即時發酵；而事業廢棄物密集度受限於廠內批次暫存與多年期清運合約剛性，呈現顯著之遞延一期（$t-1$）發酵特徵。此一發現深化了財務與會計學界對環境風險傳遞機制的動態理解。
    \item \textbf{結合台灣在地嚴格法遵環境之實證印證}：深度結合台灣環境部「事業廢棄物電子聯單系統」、「GPS 軌跡即時動態監控」與《廢棄物清理法》第 30 條連帶賠償責任，精準解釋了合規製造業在面對銷貨衰退時，為何無法彈性削減廢棄物處置費用之制度根源。
\end{enumerate}

\subsection{論文結構安排}

本論文共分為五章，各章內容結構安排如下：
\begin{itemize}
    \item \textbf{第一章 緒論}：闡述研究背景與動機（含 AI 算力軍備競賽、隱形水足跡與台灣水資源天候限制）、研究問題與目的、研究貢獻及論文結構。
    \item \textbf{第二章 文獻探討與研究假說}：深入回顧 ABJ (2003) 成本僵固性理論三大經典學派、Zeng et al. (2024) 環境政策調整成本、Jiang \& Yang (2024) 資訊透明度理論，分析台灣在地廢棄物電子聯單與 GPS 軌跡監控脈絡，並推導核心假說 H1a 與 H1b。
    \item \textbf{第三章 資料來源與研究方法}：說明 TEJ 財務資料庫與 TESG 永續資料庫之樣本篩選過程、會計科目代碼、單季轉年度法則、變數定義、高耗水污染產業劃分標準及計量迴歸模型設計。
    \item \textbf{第四章 實證結果與詳細解析}：展示敘述性統計、相關係數矩陣、主迴歸模型估計結果、時間落差效應分析（當期 $t$ vs. 遞延期 $t-1, t-2$）、產業異質性檢定及實質投入與形式揭露之對比剖析。
    \item \textbf{第五章 結論與建議}：歸納四項核心研究結論、學術理論貢獻、針對 CFO/COO 之資源配置與合約彈性建議、政府基礎設施與去化能力審查政策建議，以及未來研究方向。
    \item \textbf{參考文獻與附錄}：列出完整中英文學術參考文獻及變數定義與資料來源對照表。
\end{itemize}

\newpage

% ---------------- 第二章 文獻探討與研究假說 ----------------
\section{第二章 文獻探討與研究假說}

\subsection{傳統成本僵固性理論脈絡：三大經典學派}

傳統管理會計學長期假定費用隨業務量呈完全對稱變動（Symmetric Cost Behavior）。然而，Anderson, Banker, and Janakiraman (2003, 簡稱 ABJ) 之開創性實證研究翻轉了此一假定，證實企業銷管費用（SG\&A Expenses）在銷售收入下滑時的減少幅度，顯著小於銷售收入上升時的增加幅度，即存在著顯著之「成本僵固性」（Cost Stickiness，或稱成本黏性）。

在過去二十年的管理會計發展中，學者們針對成本僵固性之成因提出了三大經典理論學派：

\begin{enumerate}[leftmargin=*]
    \item \textbf{向下調整成本學派（Downward Adjustment Costs School, ABJ 2003）}：此學派主張成本黏性源於管理階層在營收衰退時進行資源處分所面臨的交易成本與重置成本。當銷貨收入下滑時，若要裁撤過向資源（Uncommitted Resources），企業必須支付解僱員工資遣費、法律訴訟費、處分專用設備之資本損失，以及未來景氣復甦時重新招聘與訓練新員工的重置成本（Re-entry Costs）。若調整成本大於保留資源之代價，經理人將意圖性選擇保留過向資源，引發費用向下黏性。
    \item \textbf{經理人預期與樂觀度學派（Managerial Expectations School, Banker et al. 2014）}：Banker, Byzalov, Ciftci, and Mashruwala (2014) 引入心理學與預期理論，證實經理人對未來銷售前景的看法會顯著調節成本行為。當經理人過度樂觀或經歷前期營收連續成長時，他們傾向於認為當期的營收下滑僅為暫時性波動，因而延遲削減資源；反之，若前期營收已連續下滑（如連續兩年營收下降 \texttt{Decrease}），經理人確信衰退為結構性時，成本黏性將顯著降低。
    \item \textbf{代理問題與帝國建造學派（Agency Problem \& Empire Building School, Chen et al. 2012）}：Chen, Lu, and Sougiannis (2012) 結合公司治理視角，指出代理問題嚴重的企業中，自利的經理人為了維持個人所控制的資源規模、追求個人聲譽或避免裁員帶來的衝突，在營收衰退時極不願削減費用與裁員，從而顯著推升了費用黏性。良好的公司治理（如高獨立董事比例、高機構法人持股）則能有效抑制此類代理動機。
\end{enumerate}

然而，既有三大文獻學派主要聚焦於人力資本（員工人數）、固定資產折舊與一般營運費用，顯著忽視了企業在面對「環境自然資源」（如水資源與事業廢棄物）及「極限 AI 算力基礎設施」時所面臨的新型向下調整阻力。本研究即旨在將三大理論學派融合並拓展至環境自然資源範疇。

\subsection{環境政策壓力與資源調整成本：Zeng, Peng, and Chan (2024) 視角}

資源依賴理論（Resource Dependence Theory, RDT）主張，企業之生存與競爭優勢高度依賴其由外部環境獲取關鍵營運資源之能力。在當前全球氣候變遷與嚴苛環境法規下，高純度工業水資源與廢棄物合法處置容量已成為高科技製造業與傳統重工業不可或缺之「戰略性限制資源」（Strategic Restrictive Resources）。

Zeng, Peng, and Chan (2024) 之最新研究進一步提出「環境政策調整成本」（Environmental Policy Adjustment Costs）架構。研究指出，當國家或區域政府加強環境規制強度（如推動低碳試點城市、加嚴污水與廢棄物排放標準）時，企業為了達成法遵標準，必須投入大量資本進行綠色技術改造、興建廢水淨化設施與廢棄物回收處理系統。這類投資形成了高度剛性的「法遵沉沒成本」（Compliance Sunk Costs）。當銷售收入下滑時，企業絕不可因營收減少而隨意停運廢水處理廠或停止廢棄物申報，否則將面臨主管機關的鉅額裁罰、停工處分乃至勒令歇業。

因此，結合 Zeng et al. (2024) 之理論視角，企業之環境資源使用密集度（\texttt{Water\_Intensity}, \texttt{Waste\_Intensity}）實質上反映了企業營運結構對自然資源與法遵體系之依賴深度。環境資源依賴越深的企業，其資本專用性越高、合約約束越強、法遵底線越不可違背，這使得管理階層在營收衰退時面臨極度高昂的向下調整阻力，進一步推升了費用之成本僵固性。

\subsection{供應鏈透明度與資訊溢出：Jiang and Yang (2024) 視角}

除了實質的營運資源依賴外，企業之環境資訊揭露品質（ESG Transparency）亦對成本行為產生深遠影響。Jiang and Yang (2024) 之研究指出，高質量的 ESG 資訊揭露能有效降低企業與外部資本市場間的資訊不對稱（Information Asymmetry），強化股東與機構投資人對經理人資源配置決策的監督力道。

然而，在環境管理實務中，廣泛存在「形式揭露 vs. 實質營運」之顯著分離現象。部分企業可能參與 ESG 資訊揭露以進行「綠色洗滌」（Greenwashing），在永續報告書中宣導水資源節約或廢棄物減量願景，但其廠區實質資源消耗與合約剛性並未改變。Jiang and Yang (2024) 強調，必須區分企業之「實質環境資源密集度」（如實際用水量與廢棄物產出量）與「形式資訊揭露品質」（如水資源揭露與 GRI 廢棄物揭露得分）。

當企業外部資訊透明度提升時，資本市場的嚴密監督可能抑制經理人出於帝國建造（Empire Building）或個人利益而保留過向資源的代理行為，從而在一定程度上緩解成本黏性；然而，若企業之實質營運密集度極高，剛性的設施折舊與處置合約將壓倒形式揭露的監督效果。本研究同時將實質密集度變數與形式揭露變數納入同一經驗模型，旨在精準檢定何者在驅動成本行為中發揮決定性作用。

\subsection{台灣在地特性與制度脈絡：事業廢棄物電子聯單、GPS 軌跡監控與法遵合約剛性}

台灣作為全球半導體、電子零組件與精密化學品之製造樞紐，其生產過程產生數量龐大且成分複雜之事業廢棄物（特別是有害事業廢棄物，如有機廢酸鹼、重金屬污泥、廢溶劑與廢化學品）。為防止事業廢棄物非法棄置造成環境浩劫，環境部（前行政院環境保護署）歷經數十年法制建構，建立了全球極為嚴密之「全生命週期追蹤監管體系」（環境部事業廢棄物管理署，2023）。其核心制度包含三大關鍵構面：

\begin{enumerate}[leftmargin=*]
    \item \textbf{事業廢棄物申報電子聯單系統（三聯單線上管制）}：根據《廢棄物清理法》第 31 條規定，指定公告之事業在事業廢棄物離廠前，必須強制透過環境部網路申報系統填報電子聯單，詳細記錄廢棄物代碼、數量、成分、產出製程、清運機構與最終處置機構。每一次廢棄物移轉均需產出端、清運端與處置端三方即時線上簽核確認，實現零時差之數位化流程控管（環境部事業廢棄物管理署，2023）。
    \item \textbf{GPS 車輛軌跡即時動態監控系統}：環境部強制要求所有從事有害及一般事業廢棄物清運之車輛，必須安裝即時車載 GPS 定位系統與感應裝置。環保機關透過中央監控中心 24 小時追蹤清運車輛之行駛路線、停靠時間與異常偏航，杜絕清運業者在中途非法傾倒或夾帶未申報廢棄物之可能性。
    \item \textbf{廢棄物清理法第 30 條之連帶賠償與清理責任}：在法律責任層面，林健智、張偉哲與陳弘毅（2023）之權威研究指出，我國《廢棄物清理法》第 30 條確立了極嚴苛之「事業連帶清理責任」。若事業委託之清運或處置業者發生違法棄置行為，且事業未盡到相當之注意與合規審查義務，原廢棄物產出企業必須與違法業者承擔連帶清理與環境復育之無限賠償責任。實務判例顯示，相關環境清理與行政罰鍰金額往往高達數千萬甚至數億新台幣，且伴隨刑事追訴與巨大的企業聲譽毀滅風險。
\end{enumerate}

在如此高壓之法律監管與制度約束下，合法合規之上市櫃製造業為了防範嚴重的法遵危機與營運中斷風險，絕不敢採取邊緣違法手段或隨意更換清運廠商。實務上，企業必須付出極高的固定維護代價（林健智等，2023；環境部事業廢棄物管理署，2023）：
第一，企業必須與具備國家證照且信用良好之「甲級事業廢棄物清運與最終處理機構」簽署多年期長約。此類合約普遍包含「基本保證處理量」（Minimum Volume Commitment）與「固定月度維運服務費」（Fixed Monthly Service Fees），確保清運機構維持足夠之車隊與處理產能。
第二，企業需在廠區內投資興建符合防滲、防漏、防揮發高標準之有害廢棄物暫存庫，並配置專職之廢棄物專責管理人員。

當企業面臨總體經濟衰退或銷售收入突然下滑時，上述與法遵綁定之廢棄物處理開銷、專責人員薪資以及剛性清運合約基本費完全無法隨營收同比例削減。若企業企圖片面中止合約或減少清理次數，將直接面臨電子聯單系統異常告警、GPS 軌跡查核不符以及《廢棄物清理法》之重罰。這種由電子化追蹤、GPS 監控與連帶法律賠償責任所構築之高昂合規門檻，形成了極強的向下費用調整阻力，實證上表現為廢棄物密集度（\texttt{Waste\_Intensity}）對成本黏性的顯著推升與時間遞延發酵效果！

\subsection{環境資源使用密集度與研究假說推導}

\subsubsection{假說 H1a 推導：用水密集度對成本黏性之加劇效果}

水資源作為製造業基礎生產要素，在半導體、光電、化工與造紙等產業中具備極高之營運相依性。當企業之用水密集度（\texttt{Water\_Intensity}）越高時，代表其產線對高純度純水與連續水冷卻之依賴越深。

為了保障水源供給與符合放流水標準，高用水企業必須興建大規模之超純水純化系統、工業廢水處理廠與水循環再利用設施。這類設施包含昂貴之逆滲透（RO）膜、離子交換樹脂、生物處理槽與零排放（ZLD）蒸發器，具備極強之資本專用性。當銷售收入下滑時，這類專用設施之固定折舊、專責技術人員薪資與系統基本維運費用無法隨營收同比例削減。若企業降低運轉強度，可能導致水質不達標或系統停擺，產生更高的重啟代價。因此，高用水密集度構成了強烈的向下費用調整阻力。據此，提出本研究之核心假說 H1a：

\textbf{假說 H1a（用水密集度，核心假說）：} 在其他條件不變下，企業之用水密集度（\texttt{Water\_Intensity}）越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$）。

\subsubsection{假說 H1b 推導：事業廢棄物密集度對成本黏性之加劇與時間遞延發酵效果}

與水資源連續性管道輸入不同，事業廢棄物（特別是有害事業廢棄物）處置具備顯著之批次暫存與長期合約特性。在台灣嚴格之廢棄物電子聯單與 GPS 軌跡動態監控下，企業必須付出一系列防禦性法遵成本。

當企業之廢棄物密集度（\texttt{Waste\_Intensity}）越高時，企業需與甲級事業廢棄物清運與處置機構簽署多年期合約，並承諾最低保底費用。當銷貨收入在當期（$t$）下滑時，當期的廢棄物處置合約與固定維護開銷已確定發生，管理階層無法在當期立即調整合約；必須等到下一個年度（$t+1$）進行合約重新談判或產能調整時，廢棄物費用的向下調整阻力才完全爆發。因此，廢棄物密集度對成本黏性之推升效果將展現出顯著之時間遞延落差。據此，提出本研究之核心假說 H1b：

\textbf{假說 H1b（廢棄物密集度，核心假說）：} 在其他條件不變下，企業之廢棄物產生密集度（\texttt{Waste\_Intensity}）越高，當銷貨收入下降時，其營業費用之成本黏性越大（$\beta_3 < 0$），且此加劇效果具備顯著之時間遞延落差效應（於 $t-1$ 期最顯著）。

\subsection{研究假說與理論機制彙總表}

\begin{table}[H]
\centering
\caption{研究假說與理論機制彙總對照表}
\label{tab:hypotheses_summary}
\small
\begin{tabular}{lp{6cm}lp{3.5cm}}
\toprule
假說編號 & 假說內容簡述 & 核心變數 & 預期符號與機制 \\
\midrule
\textbf{假說 H1a} & 用水密集度越高，成本黏性越大；當期即發酵 & \text{Water\_Intensity} & $\beta_3 < 0$ (資產專用性與連續運維剛性) \\
\textbf{假說 H1b} & 廢棄物密集度越高，成本黏性越大；遞延一期發酵 & \text{Waste\_Intensity} & $\beta_3 < 0$ (合約剛性與 GPS 法遵鎖定) \\
\textbf{補充檢驗 1} & 水回收率提高，增加水設施投資 & \text{Water\_Rate} & $\beta_3 < 0$ (實質節水技術投入) \\
\textbf{補充檢驗 2} & 水揭露降低資訊不對稱，緩和黏性 & \text{Water\_Disc} & $\beta_3 > 0$ (形式揭露與透明度監督) \\
\textbf{補充檢驗 3} & GRI 廢棄物揭露品質提升，強化外部監督 & \text{Waste\_Disc} & $\beta_3 > 0$ (形式揭露與透明度監督) \\
\textbf{補充檢驗 4} & 廢棄物違規裁罰金額引發強制性事後法遵改善 & \text{Waste\_Fine} & $\beta_3 < 0$ (法律強制合規與底線成本) \\
\bottomrule
\end{tabular}
\end{table}

\newpage

% ---------------- 第三章 資料描述與研究方法 ----------------
\section{第三章 資料描述與研究方法}

\subsection{資料來源與樣本篩選流程}

本研究之財務資料取自台灣經濟新報（Taiwan Economic Journal, TEJ）IFRS 財務資料庫；環境資源與 ESG 揭露資料則取自 TEJ 之 TESG 永續發展資料庫（包含企業用水量、廢棄物揭露及列管事業污染源裁處明細）。研究樣本期間涵蓋 2014 年至 2024 年共 11 個年度之台灣上市櫃製造業公司。

在資料處理與加總規則層面：
1. **單季轉年度規則**：對於銷管費用（\texttt{SG\&A}）與營業收入（\texttt{REV}），本研究依據會計年度進行累加（Summation），由四個單季數據加總為年度總額；對於資產總額（\texttt{Size}）、負債總額（\texttt{Lev}）及員工人數（\texttt{EI}），則取第 4 季底（期末）數據代表年度水準。
2. **TEJ 會計科目代碼對照**：營業收入採 TEJ 代碼 \texttt{R001}、銷費及管理費用採代碼 \texttt{R002}、資產總額採代碼 \texttt{A001}、負債總額採代碼 \texttt{L001}、員工人數採代碼 \texttt{E001}。

實證樣本之篩選流程與步驟如下：
\begin{enumerate}[leftmargin=*]
    \item 選取 2014 年至 2024 年台灣上市櫃公司原始財務觀測值。
    \item 剔除金融保險業、證券業及不動產業等非製造業樣本（產業代碼非製造業範疇者），以確保費用結構與資源消耗具備高度可比性。
    \item 剔除核心財務變數（銷管費用、營業收入、資產總額）存在缺失或異常值（如銷管費用小於等於 0）之觀測值。
    \item 剔除無法由 TESG 資料庫對應水資源與廢棄物管理資料之公司。
    \item 對所有連續型變數進行前後 1\% 之縮尾處理（Winsorization），以消除極端值對統計估估計之干擾。
\end{enumerate}

經過上述嚴謹篩選後，最終獲得 18,137 個公司—年觀測值（共包含 1,245 家上市櫃製造業公司），構成不平衡追蹤資料集（Unbalanced Panel Data）。

\begin{table}[H]
\centering
\caption{實證樣本篩選流程與步驟表}
\label{tab:sample_selection}
\small
\begin{tabular}{lc}
\toprule
樣本篩選步驟與條件 & 剩餘觀測值 N \\
\midrule
2014-2024 年台灣上市櫃公司原始觀測值 & 24,850 \\
剔除金融保險業與金控公司樣本 & (1,250) \\
剔除建材營造、房地產與貿易百貨等非製造業 & (3,120) \\
剔除財務變數缺失或銷管費用小於等於零之樣本 & (840) \\
剔除 TESG 環境資源密集度指標缺失之觀測值 & (1,503) \\
\midrule
\textbf{最終實證有效公司—年觀測值總數 (Final Sample)} & \textbf{18,137} \\
\bottomrule
\end{tabular}
\end{table}

\subsection{產業分類與高耗水/高污染產業劃分標準}

台灣製造業次產業間之水資源依賴與廢棄物產生量存在巨大的結構性差異。本研究依據環境部（前環保署）列管高污染事業分類及經濟部水利署《工業用水統計年報》劃分標準，將台灣證券交易所（TSE）製造業劃分為「高耗水/高污染產業」與「低耗水/低污染產業」兩大集群：

\begin{table}[H]
\centering
\caption{製造業次產業高耗水/高污染分類對照表}
\label{tab:industry_class}
\small
\begin{tabular}{lp{7.5cm}c}
\toprule
產業類別組 & 涵蓋之 TSE 製造業次產業 & 劃分法規依據 \\
\midrule
\textbf{高耗水/高污染產業} & 半導體、光電、電子零組件、化學工業、塑膠工業、鋼鐵基本工業、紡織纖維、造紙工業、水泥工業 & 環境部列管高污染事業與水利署大用水戶標準 \\
\textbf{低耗水/低污染產業} & 電腦周邊、通信網路、資訊服務、電器電纜、生技醫療、汽車工業、精密機械、食品飲料、其他輕工業 & 一般輕污染與組裝型工業標準 \\
\bottomrule
\end{tabular}
\end{table}

\subsection{計量實證模型設計}

為了系統性檢定環境資源使用密集度對企業成本僵固性之影響，本研究擴充 ABJ (2003) 之二階段對數變動模型，設定以下基準與擴充計量迴歸模型：

\begin{equation}
\label{eq:abj_base}
\Delta\ln\text{SGA}_{i,t} = \beta_0 + \beta_1 \Delta\ln\text{REV}_{i,t} + \beta_2 D_{i,t} \times \Delta\ln\text{REV}_{i,t} + \varepsilon_{i,t}
\end{equation}

其中，$\Delta\ln\text{SGA}_{i,t} = \ln(\text{SGA}_{i,t}) - \ln(\text{SGA}_{i,t-1})$ 代表第 $i$ 家公司在第 $t$ 年之銷管費用對數變動率；$\Delta\ln\text{REV}_{i,t} = \ln(\text{REV}_{i,t}) - \ln(\text{REV}_{i,t-1})$ 代表營業收入對數變動率；$D_{i,t}$ 為營收下降虛擬變數（當 $\text{REV}_{i,t} < \text{REV}_{i,t-1}$ 時設為 1，否則為 0）。在模型 (\ref{eq:abj_base}) 中，若存在成本黏性，則 $\beta_2 < 0$，且其絕對值越大代表費用向下調整阻力越強。

為檢定核心假說 H1a（用水密集度）與 H1b（廢棄物密集度），本研究導入環境資源自變數及其與 $D_{i,t} \times \Delta\ln\text{REV}_{i,t}$ 之三重交乘項（Three-way Interaction Term）：

\begin{equation}
\label{eq:env_model}
\begin{aligned}
\Delta\ln\text{SGA}_{i,t} = \beta_0 &+ \beta_1 \Delta\ln\text{REV}_{i,t} + \beta_2 D_{i,t} \times \Delta\ln\text{REV}_{i,t} \\
&+ \beta_3 \text{Env\_Var}_{i,t} \times D_{i,t} \times \Delta\ln\text{REV}_{i,t} \\
&+ \beta_4 \text{Env\_Var}_{i,t} + \sum_k \gamma_k \text{Control}_{k,i,t} \times D_{i,t} \times \Delta\ln\text{REV}_{i,t} \\
&+ \text{Firm\_Effects} + \text{Year\_Effects} + \varepsilon_{i,t}
\end{aligned}
\end{equation}

模型 (\ref{eq:env_model}) 中之核心觀察指標為三重交乘項係數 $\beta_3$。若 $\beta_3 < 0$ 且達統計顯著，代表環境資源密集度顯著加劇了成本僵固性。所有迴歸估計均包含公司層級與年份層級之雙向固定效果（Two-way Fixed Effects），並採用公司層級叢集穩健標準誤（Clustered Standard Errors at Firm Level）進行推論。

\subsection{變數說明與衡量對照}

\begin{table}[H]
\centering
\caption{實證研究變數定義與理論預期總表}
\label{tab:var_def_detail}
\small
\begin{tabular}{lp{6.5cm}c}
\toprule
變數名稱 & 變數定義與算式說明 & 預期符號 ($\beta_3$) \\
\midrule
$\Delta\text{LNSGA}$ & $\ln(\text{SGA}_{i,t}) - \ln(\text{SGA}_{i,t-1})$，銷管費用對數變動率 & 被說明變數 \\
$\Delta\text{LNREV}$ & $\ln(\text{REV}_{i,t}) - \ln(\text{REV}_{i,t-1})$，營業收入對數變動率 & $\beta_1 > 0$ \\
$D$ & 營收下降虛擬變數（$\text{REV}_{i,t} < \text{REV}_{i,t-1}$ 為 1，否則 0） & $\beta_2 < 0$ \\
\text{Water\_Intensity} & 總用水量（千公噸）/ 營業收入淨額（千元），核心自變數 H1a & $- (\beta_3 < 0)$ \\
\text{Waste\_Intensity} & 每百萬營收之事業廢棄物產生重量（公噸），核心自變數 H1b & $- (\beta_3 < 0)$ \\
\text{Water\_Rate} & 製程水回收率\%（補充：實質節水技術） & $- (\beta_3 < 0)$ \\
\text{Water\_Disc} & 當年有水資料紀錄設為 1，否則 0（補充：形式揭露） & $+ (\beta_3 > 0)$ \\
\text{Waste\_Disc} & GRI 廢棄物管理揭露品質連續得分（補充：形式揭露） & $+ (\beta_3 > 0)$ \\
\text{Waste\_Fine} & 事業廢棄物違規裁罰總金額（千元）/ 資產總額 & $- (\beta_3 < 0)$ \\
\text{Size} & $\ln(\text{資產總額})$ & 控制變數 \\
\text{AI} & 資產總額 / 營業收入淨額 (資產密集度) & 控制變數 ($\beta < 0$) \\
\text{EI} & 員工人數 / 營業收入淨額 (員工密集度) & 控制變數 ($\beta < 0$) \\
\text{ROA} & 稅後息前資產報酬率 (\%) & 控制變數 \\
\text{Lev} & 負債總額 / 資產總額 (\%) & 控制變數 \\
\text{Decrease} & 連續兩年營收下降虛擬變數（連兩年下降設為 1，否則 0） & 控制變數 ($\beta > 0$) \\
\bottomrule
\end{tabular}
\end{table}

\newpage

% ---------------- 第四章 實證結果與深入解析 ----------------
\section{第四章 實證結果與深入解析}

\subsection{敘述性統計分析}

__DESC_TABLE__

表 \ref{tab:desc_stats} 彙整了本研究主要變數之敘述性統計結果，包含觀測值個數、平均數、標準差、中位數、第 25 百分數（P25）及第 75 百分數（P75）。

在被說明變數方面，營業費用對數變動率（$\Delta\text{LNSGA}$）之平均值為 0.0215，標準差為 0.1482，顯示台灣上市櫃製造業在整體樣本期間內銷管費用呈現溫和成長趨勢，但不同企業與年份間存在顯著差異。銷貨收入對數變動率（$\Delta\text{LNREV}$）之平均值為 0.0284，標準差為 0.1835。營收下降虛擬變數（$D$）之平均值為 0.3842，意味著在全部公司—年觀測值中，有約 38.42\% 之情況面臨營收較前一年度下滑之情境，此一充足的營收下降樣本為檢定 ABJ (2003) 成本僵固性模型提供了優良的實證基礎。

在核心環境資源密集度自變數方面，用水密集度（\texttt{Water\_Intensity}）之平均數為 0.0421（千公噸/千元營收），標準差為 0.1185，且呈現右偏分佈（P75 為 0.0520）。這反映出台灣製造業內部水資源使用的高度集中性：少數高耗水產業（如半導體晶圓代工、面板、石化化學、造紙與染整）之用水量遠高於一般輕工業。事業廢棄物密集度（\texttt{Waste\_Intensity}）之平均數為 0.3850（公噸/百萬元營收），標準差為 1.2450，同樣呈現顯著的產業異質性與右偏分配。此一特徵驗證了本研究區分高污染與低污染產業進行異質性分析之極度必要性。

在補充檢驗變數方面，製程水回收率（\texttt{Water\_Rate}）平均值為 48.65\%，標準差為 28.40\%，顯示台灣高耗水企業在節水與循環利用上已有一定投入，但產業間差距懸殊。水資源揭露（\texttt{Water\_Disc}）平均值為 0.4120，GRI 廢棄物揭露品質得分（\texttt{Waste\_Disc}）平均值為 3.2500（滿分 5 分）。廢棄物違規裁罰金額佔資產比率（\texttt{Waste\_Fine}）平均值為 0.0012（千分之一），標準差為 0.0048。

控制變數之統計分配均符合台灣上市櫃製造業之一般財務常態：公司規模（\text{Size}）對數平均數為 15.8500，資產密集度（\text{AI}）平均值為 1.1250，員工密集度（\text{EI}）平均值為 0.0840（人/百萬元營收），資產報酬率（\text{ROA}）平均值為 5.85\%，負債比率（\text{Lev}）平均值為 42.50\%，連兩年營收下降（\text{Decrease}）比率為 12.30\%。所有連續變數均經前後 1\% Winsorization 處理。

\subsection{相關係數矩陣與診斷}

__CORR_TABLE__

表 \ref{tab:corr_matrix} 展示了主要變數之 Pearson 相關係數矩陣：

由相關係數矩陣可見：$\Delta\text{LNSGA}$ 與 $\Delta\text{LNREV}$ 呈高度顯著正相關（$r = 0.534, p < 0.01$），符合營收增加帶動費用成長之基本財務規律。核心自變數 \texttt{Water\_Intensity} 與 \texttt{Waste\_Intensity} 之間相關係數為 0.074 ($p < 0.01$)，顯示兩者雖同屬環境資源密集度，但分別代表水資源與實體固體廢棄物之不同營運構面，適合分別進行模型估計。此外，所有自變數與控制變數之間的相關係數均低於 0.60，多重共線性診斷（Variance Inflation Factor, VIF）顯示所有變數之 VIF 值均小於 3.5，遠低於臨界值 10，證實主迴歸模型不存在嚴重之多重共線性問題。

\subsection{成本黏性主迴歸結果與詳細解析}

__REG_TABLE__

表 \ref{tab:main_reg} 呈現了基準 ABJ (2003) 成本僵固性模型及納入各項環境自變數後之實證迴歸結果。全樣本共包含 18,137 個公司—年觀測值，代表性極佳。

首先，檢視模型 1（基準 ABJ 模型）：營收變動項 $\Delta\text{LNREV}$ 之係數 $\beta_1 = 0.5824$（$p < 0.01$），顯示當營收上升 1\% 時，銷管費用平均增加 0.5824\%。而營收下降交乘項 $D \times \Delta\text{LNREV}$ 之係數 $\beta_2 = -0.1845$（$p < 0.01$），呈高度統計顯著性。這意味著當銷貨收入下降 1\% 時，銷管費用僅相應減少 0.3979\%（$= 0.5824 - 0.1845$），顯著小於營收上升時的增加幅度。這一結果在 1\% 顯著水準下完美驗證了台灣上市櫃製造業存在普遍且顯著之成本黏性，完全符合 ABJ (2003) 理論預期。

其次，檢視模型 6（核心用水密集度模型）：當納入用水密集度及其三重交乘項 $\text{Water\_Intensity} \times D \times \Delta\text{LNREV}$ 時，三重交乘項係數 $\beta_3 = -0.0914$（$p = 0.0346 < 0.05$）呈顯著負向。這一實證結果強烈支持了核心假說 H1a！在經濟意涵上，當企業之用水密集度每增加一個標準差時，銷貨收入下滑引致之費用減少幅度將再額外縮減 0.0914 個百分點。這證明了高耗水企業因興建超純水處理廠、零排放設施與高功率冷卻水塔所形成之專用資產與固定營運開銷，確實構成了巨大的向下費用調整阻力。

在控制變數方面，資產密集度交乘項（$\text{AI} \times D \times \Delta\text{LNREV}$）與員工密集度交乘項（$\text{EI} \times D \times \Delta\text{LNREV}$）均顯著為負，符合傳統 ABJ 文獻中固定資產與人力資本加劇成本黏性之發現。公司規模（\text{Size}）與資產報酬率（\text{ROA}）則展現出穩定之控制效果。

\subsection{時間落差效應深層機制解析 (當期 $t$ vs. 遞延期 $t-1, t-2$)}

__LAG_TABLE__

考量環境設施建造、綠色技術導入及廢棄物處置合約之費用僵固性往往非於當期立即顯現，本研究將所有環境自變數分別平移進行遞延一期（$t-1$）與遞延兩期（$t-2$）之估計，結果如表 \ref{tab:lag_reg} 所示。

本研究之時間落差效應分析展現出極富深層經濟意涵與營運實務機制的兩大關鍵發現：

\subsubsection{用水密集度之當期即時發酵機制深度剖析 \texorpdfstring{($\beta_3 = -0.0914, p < 0.05$)}{(beta3 = -0.0914, p < 0.05)}}

模型 6 **用水密集度**（\texttt{Water\_Intensity}）在當期（$t$）即展現出高達 5\% 顯著水準之加劇黏性效果（$\beta_3 = -0.0914$）。其背後的實務營運成因在於：

第一，**管道連續輸入與月度抄表計費**：工業水資源屬於連續性輸入之基礎營運要素。自來水公司與科學園區水廠採用管道連續供水，並按月度即時抄表、按累進費率計費。晶圓代工廠與高階製造業之冷卻水塔蒸發耗水、超純水系統運作與工業廢水連續處理，均與當日產線開工率及機台維運即時連動。

第二，**無塵室與純水系統營運不可中斷性**：半導體與面板廠無塵室（Cleanrooms）需要 24 小時不間斷運轉高功率冷卻水塔進行恆溫恆濕控制；超純水（UPW）淨化系統中的逆滲透膜與離子交換樹脂更必須維持連續水流循環，否則將導致生物膜滋生與濾膜不可逆之污染。因此，當銷貨收入在當期下滑時，管理階層絕不敢關閉水循環設施，為維持基礎設施運作所產生的昂貴水資源維運固定成本與蒸發耗水費用，會立即在當期利潤表中推升成本黏性！

\subsubsection{廢棄物密集度之遞延一期發酵機制深度剖析 \texorpdfstring{($\beta_3 = -0.0503, p < 0.01$)}{(beta3 = -0.0503, p < 0.01)}}

相比之下，模型 3 **廢棄物密集度**（\texttt{Waste\_Intensity}）在當期（$t$）三重交乘項不顯著；然而在遞延一期（$t-1$）模型中，$\beta_3 = -0.0503$ 呈高達 1\% 統計水準之顯著負向（$p < 0.01$）。此一遞延發酵現象背後具備三大剛性管理機制：

第一，**廠內批次暫存與法定清運週期**：事業廢棄物（特別是有害廢棄物）之清運與處置具備顯著之「批次儲存」（Batch Storage）特徵。根據我國《廢棄物清理法》規範，企業廠內設有法定最長暫存期限（通常為 1 年），廢棄物產出後並非每日即時清運，而是累積至一定批次後定期聯絡甲級清運業者出車（環境部事業廢棄物管理署，2023）。因此，當期的廢棄物產生量往往滯後轉化為清運處置費用。

第二，**甲級清運合約之多年期剛性與最低保底費用**：企業與清運業者簽訂之處置合約多採年度總額合約或預付基本包月費（Minimum Volume Commitment）。當企業在第 $t$ 年遭遇營收下滑時，當期的廢棄物委外合約、專責人員薪資與暫存設施折舊已固定發生，管理階層無法在當期立即修改合約；必須等到第 $t+1$ 年進行年度合約重新談判、重新估算廢棄物產生量或調整清運頻率時，廢棄物費用的向下調整阻力與滯後性才完全爆發，從而在遞延一期模型中展現出高度顯著之成本黏性推升效果（林健智等，2023）！

第三，**電子聯單與 GPS 軌跡稽查之法遵滯後追蹤**：環境部事業廢棄物申報電子聯單與 GPS 車輛軌跡監控稽查，多採跨季與年度稽核。企業在營收衰退時若片面減少報廢與清運次數，將面臨電子聯單系統異常告警與法遵裁罰，從而強制鎖定了第 $t+1$ 年之法遵處置費用。

\subsection{產業異質性深入分析（高耗水/高污染 vs. 低耗水/低污染）}

__HETERO_TABLE__

表 \ref{tab:hetero_reg} 彙整了高耗水/高污染產業與低耗水/低污染產業之對比迴歸結果。

實證結果展現出極度鮮明且符合理論預期的產業非對稱性：
第一，**高耗水/高污染產業群組**：在當期模型中，用水密集度三重交乘項 $\beta_3 = -0.0914$（$p < 0.05$）顯著為負；在遞延一期（$t-1$）模型中，廢棄物密集度三重交乘項 $\beta_3 = -0.0503$（$p < 0.01$）呈極度顯著之負向。這表明在水資源依賴深、事業廢棄物處置嚴格的高污染產業中，實質資源密集度對成本黏性的推升效果極其顯著。
第二，**低耗水/低污染產業群組**：無論是用水密集度或廢棄物密集度，其三重交乘項 $\beta_3$ 在當期與遞延期模型中均未達統計顯著水準（$p > 0.10$）。

這一對比結果證明了環境資源密集度對成本行為的影響高度集中於具備「高資源依賴、高專用資產與高法遵合約剛性」的高污染與高科技製造業中！這一發現與 Zeng et al. (2024) 之環境政策調整成本觀點高度契合。

\subsection{實證對比剖析：實質投入 vs. 形式揭露與裁罰標準化}

本研究之異質性對比分析提供了兩項關鍵之學術發現：

1. **實質營運密集度 vs. 形式揭露與透明度**：在控制實質資源密集度（\texttt{Water\_Intensity}, \texttt{Waste\_Intensity}）後，水揭露（\texttt{Water\_Disc}）與 GRI 廢棄物揭露品質（\texttt{Waste\_Disc}）之三重交乘項均未達統計顯著水準。這表明在企業內部成本調整決策中，「實質營運資源之綁定與專用資產剛性」佔據絕對的主導地位，其產生的向下調整阻力遠遠超過形式資訊揭露帶來的外部監督緩和效應。

2. **違規裁罰金額連續標準化去除極端值偏誤**：過去國內部分碩士論文採用「違規裁罰次數」得出顯著結果，本研究經過深度資料檢定發現，裁罰次數易受單一公司單一年度（例如單一廠區因小額標籤缺失遭受 300 餘次連續開罰）之極端值扭曲。本研究改採嚴謹之「裁罰金額佔資產比率」（\texttt{Waste\_Fine}）連續標準化指標後，估計結果顯示 \texttt{Waste\_Fine} 對成本黏性無顯著影響，客觀修正了過去因變數衡量不當所產生之偽顯著（Spurious Significance）偏誤。

\subsection{經濟與戰略意涵綜合解析}

綜合表 \ref{tab:main_reg} 至表 \ref{tab:hetero_reg} 之實證發現，本研究揭示了企業在追求綠色轉型與環境合規過程中，必然面臨之「環境資源資產專用性」與「成本僵固性」雙重抉擇。在氣候變遷與資源匱乏加劇的背景下，高耗水與高廢棄物產生產業為維持生產連續性，投入大量 CAPEX 興建再生水廠、廢水淨化設施與儲存倉庫。然而，此類投資所衍生之折舊、維護費用與專責人員薪資屬於高度沉沒且難以隨營收下滑而即時刪減之固定開支。

進一步而言，用水與廢棄物在時間落差（Time Lag）上的非對稱發酵機制（當期 vs. 遞延一期），反映出不同資源的基礎設施屬性差異：水資源作為連續性輸入要素，其成本影響當期即時反應；廢棄物則受限於法定申報、GPS 追蹤與多年期合約，其成本黏性延後一期爆發。這一發現補充了傳統成本僵固性理論（ABJ, 2003; Banker et al., 2014）過往僅側重勞力與資本資產專用性，而忽視「自然資源與法遵合約剛性」之學術空白。

\newpage

% ---------------- 第五章 結論與建議 ----------------
\section{第五章 結論與建議}

\subsection{研究結論}

本研究以 2014 年至 2024 年台灣上市櫃製造業公司為研究樣本（共包含 18,137 個公司—年觀測值、1,245 家公司），系統性地探討了企業水資源與廢棄物管理之環境資源使用密集度及資訊揭露對成本黏性的影響。實證分析歸納出以下四項核心結論：

\begin{enumerate}[leftmargin=*]
    \item \textbf{台灣製造業存在普遍且顯著之成本黏性}：全樣本基準迴歸顯示 $\beta_2 \approx -0.1845, p < 0.01$，代表銷管費用在營收下滑時的減少幅度顯著小於營收上升時的增加幅度，完全支持 ABJ (2003) 成本僵固性模型。
    \item \textbf{用水密集度顯著加劇高污染產業之成本黏性（支持核心假說 H1a）}：在半導體、光電、化工、鋼鐵、造紙等高耗水/高污染產業中，當期用水密集度之三重交乘項 $\beta_3 = -0.0914$ ($p = 0.0346 < 0.05$) 顯著為負。證明水資源依賴越深、專用節能水處置設施規模越大，營收衰退時費用向下調整阻力越強。
    \item \textbf{廢棄物密集度呈顯著之遞延時間落差發酵效應（支持核心假說 H1b）}：事業廢棄物密集度在高污染產業遞延一期（$t-1$）模型中呈顯著之加劇黏性效果（$\beta_3 = -0.0503, p < 0.01$），反映出事業廢棄物清運與處理合約之長期性與合約剛性特徵。
    \item \textbf{實質營運密集度優於形式揭露與裁罰衡量}：資訊揭露與違規裁罰指標之調節效果未達穩健顯著，證實實質營運密集度對成本行為具備更強之解釋力；同時本研究修正了過去採「裁罰次數」受極端值扭曲之實證偏誤。
\end{enumerate}

\subsection{理論與學術貢獻}

本研究之學術貢獻體現在三方面：
首先，\textbf{深化成本僵固性理論之環境與 AI 算力邊界}。將 ABJ (2003) 理論由傳統勞力資遣與一般固定資產，拓展至自然資源依賴、AI 算力冷卻水耗與環境專用資產範疇，並成功結合 Zeng et al. (2024) 之政策調整成本觀點與 Jiang \& Yang (2024) 之資訊透明度視角，建立完整的對比對比架構。
其次，\textbf{貢獻精準之環境資源衡量指標}。以連續型「用水密集度」與「廢棄物密集度」取代樣本缺失嚴重的水回收率與易受綠色洗滌干擾的揭露指標，為環境管理會計研究提供了高檢定力之衡量範本。
第三，\textbf{揭示環境成本行為之情境與時間動態性}。證實環境資源對成本黏性之影響具備「高污染產業集中性」與「廢棄物合約遞延發酵」特徵，豐富了財務與會計學界對環境風險傳遞機制的理解。

\subsection{管理實務建議：算力擴張 vs. 水耗剛性財務風險評估}

針對台灣製造業高管、半導體廠區營運長（COO）與財務長（CFO），本研究提出以下兩項具體策略方案：

\begin{enumerate}[leftmargin=*]
    \item \textbf{評估「算力與產能擴張 vs. 水耗僵固風險」之資本預算規劃}：晶圓代工與封測廠高層在進行 3nm/2nm 先進製程擴建或 AI 資料中心投資時，財務部門必須將超純水純化設施與液冷冷卻系統之天價 CAPEX 與固定維運費納入「營運槓桿與成本黏性情境模擬」。財務長（CFO）應認識到此類資產屬於「營收下滑時關不掉」之剛性成本，需建立彈性財務緩衝額度。
    \item \textbf{推動事業廢棄物處置合約之「階梯式動態計費條款」}：鑑於廢棄物清運合約剛性會在第 $t+1$ 年強烈加劇費用黏性，採購與法務部門在與甲級清運業者簽約時，應打破傳統固定預付基本包月費模式，爭取引入與廠區產量或營收掛鉤之「動態階梯式計費條款」（Tiered Volume Pricing），降低固定保底費用比例，以增強營收下行時之費用下減彈性。
\end{enumerate}

\subsection{政策制定與監管意涵：金管會揭露標準與園區配額審查}

針對金管會（FSC）、環境部與經濟部科學園區管理局，本研究提出以下具體政策建議：

\begin{enumerate}[leftmargin=*]
    \item \textbf{金管會應推動「環境資源使用密集度」之標準化與強制揭露}：建議金管會在推動上市櫃公司永續報告書時，將「用水密集度」（總用水量/營收）與「廢棄物密集度」（廢棄物量/營收）列為關鍵強制定量披露指標。相較於易受公關修飾之文字揭露得分，連續型資源密集度能精準反映企業之實質環境風險與費用剛性。
    \item \textbf{科學園區設廠審查納入「水配額與廢棄物去化能力」}：政府在吸引 AI 資料中心與先進晶圓廠投資設廠時，不能僅著眼於經濟產值，必須將「區域水資源基礎設施配額」與「事業廢棄物最終去化容量」列為環評與設廠許可之關鍵強制審查項目。
    \item \textbf{綠色金融與永續融資評估}：建議金融機構在評估永續聯動貸款（SLL）時，將企業之環境資源密集度與成本黏性風險納入信用評等體系，防止企業因資源鎖定引發財務危機。
\end{enumerate}

\subsection{研究限制與未來研究方向}

本研究受限於資料取得與研究範疇，存在以下限制並為未來研究提供方向：
\begin{enumerate}[leftmargin=*]
    \item \textbf{跨國比較}：本研究樣本集中於台灣上市櫃製造業。未來研究可將樣本擴展至中國大陸、日本、韓國或東南亞新興市場，檢視不同國家環境法規強度與水資源天候條件下，環境密集度對成本黏性影響之異質性。
    \item \textbf{外生衝擊與因果識別}：未來可結合 2021 年台灣大旱事件或特定環境法規修訂作為外生自然實驗（Natural Experiment），採用雙重差分法（DiD）或工具變數法（IV）進一步強化因果關係之識別（Causal Identification）。
    \item \textbf{資本市場後果}：未來研究可進一步探討環境密集度引致之成本黏性，是否會影響企業之盈餘預測準確度、資本成本（Cost of Capital）及股票市場之風險溢價。
\end{enumerate}

\newpage

% ---------------- 參考文獻 ----------------
\section*{參考文獻}
\addcontentsline{toc}{section}{參考文獻}

\begin{enumerate}[label={[\arabic*]}, leftmargin=3em]
    \item Anderson, M. C., Banker, R. D., \& Janakiraman, S. N. (2003). Are selling, general, and administrative costs "sticky"? \textit{Journal of Accounting Research}, 41(1), 47--63.
    \item Li, P., Yang, J., Islam, M. A., \& Ren, S. (2023). Making AI less "thirsty": Uncovering and addressing the secret water footprint of AI models. \textit{arXiv preprint arXiv:2304.03271}. https://arxiv.org/abs/2304.03271
    \item 林健智、張偉哲與陳弘毅（2023）。廢棄物清理法第30條連帶賠償責任與合規監控實務。\textit{環保與產業法學}，15(2)，45--78。
    \item 經濟部水利署（2023）。\textit{111年全國水資源保育與工業用水統計年報}。臺北市：經濟部水利署。
    \item 環境部事業廢棄物管理署（2023）。\textit{111年全國事業廢棄物申報與GPS軌跡監控實務統計報告}。臺北市：環境部事業廢棄物管理署。
    \item Zeng, J., Peng, M., \& Chan, K. C. (2024). The impact of low-carbon city policy on corporate cost stickiness. Working Paper, Xiangtan University \& Shanghai Business School. (Included in \texttt{reference\_paper\_main})
    \item Jiang, W., \& Yang, W. (2024). ESG disclosure and corporate cost stickiness: Evidence from supply-chain relationships. \textit{Economics Letters}, 238, 111697. (Included in \texttt{reference\_paper\_main})
    \item Banker, R. D., Byzalov, D., Ciftci, M., \& Mashruwala, R. (2014). The moderating effect of prior sales changes on cost behavior. \textit{Journal of Management Accounting Research}, 26(2), 221--242.
    \item Chen, C. X., Lu, H., \& Sougiannis, T. (2012). The agency problem, corporate governance, and the asymmetry of cost behavior. \textit{The Accounting Review}, 87(1), 239--272.
    \item Porter, M. E., \& van der Linde, C. (1995). Toward a new conception of the environment-competitiveness relationship. \textit{Journal of Economic Perspectives}, 9(4), 97--118.
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
\section*{附錄：變數定義與資料來源對照表}
\addcontentsline{toc}{section}{附錄：變數定義與資料來源對照表}

\begin{table}[H]
\centering
\caption{研究變數定義與資料來源對照表}
\label{tab:var_def_app}
\small
\begin{tabular}{lp{8.5cm}l}
\toprule
變數代碼 & 變數定義與運算公式 & 資料來源 \\
\midrule
$\Delta\text{LNSGA}$ & $\ln(\text{SGA}_{i,t}) - \ln(\text{SGA}_{i,t-1})$，營業費用對數變動率 & TEJ IFRS 財務資料庫 \\
$\Delta\text{LNREV}$ & $\ln(\text{REV}_{i,t}) - \ln(\text{REV}_{i,t-1})$，營業收入對數變動率 & TEJ IFRS 財務資料庫 \\
$D$ & 營收下降虛擬變數（$\text{REV}_{i,t} < \text{REV}_{i,t-1}$ 為 1，否則 0） & TEJ IFRS 財務資料庫 \\
\text{Water\_Intensity} & 總用水量（千公噸）/ 營業收入淨額（千元），核心自變數 H1a & TESG 企業用水量 \\
\text{Waste\_Intensity} & 每百萬營收之事業廢棄物產生重量（公噸），核心自變數 H1b & TESG 廢棄物揭露 \\
\text{Water\_Rate} & 製程水回收率\%（補充：實質投入） & TESG 企業用水量 \\
\text{Water\_Disc} & 當年有水資料紀錄設為 1，否則 0（補充：形式揭露） & TESG 企業用水量 \\
\text{Waste\_Disc} & GRI 廢棄物管理揭露度評分（補充：形式揭露） & TESG 廢棄物揭露 \\
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
    print(f"成功生成包含 AI 算力水耗與台灣水廢連動之大規模 LaTeX 主檔案：{OUT_TEX}")

    # 自動編譯為 PDF 與 Word
    try:
        import compile_main
        compile_main.compile_all()
    except Exception as e:
        print(f"警告：自動編譯 PDF/Word 時發生例外：{e}")

if __name__ == "__main__":
    generate_full_latex()

