import re
import os

file_path = "d:/Repositories/nsysu_finance_paper_v3/09-LaTeX論文產出與編譯/09_generate_latex.py"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Fix \\beta_3
content = content.replace(r"$\\beta_3", r"$\beta_3")

# 2. Simplify citations
replacements = {
    "（鄭紹君，2018；黃倩仙，2017；林有志等，2011）": "（林有志等，2011）",
    "（張芳瑜，2023；謝雅筑，2024）": "（Banker et al., 2014）",
    "（王惠洳，2025；顏宥盈，2024；陳思婷，2025；張宸杰，2023；邱圓雅，2020；張凱瑜，2025；劉佳熒，2022）": "（Chen et al., 2012）",
    "（許嘉芸，2026；杜雨容，2024；陳苡瑄，2024）": "（許嘉芸，2026）",
    "（楊博勝，2026；蘇姿云，2026；李亞峮，2026；劉翠山，2026；吳東蓄，2026；戴郁樺，2026；楊宜蓁，2026；楊珆佳，2026；富得運，2026；毛祚祥，2026；陳俊佑，2026；張湘泓，2025）": "（楊博勝，2026）",
    "（張岑華，2026；張筱婕，2026）": "（張岑華，2026）",
    "（周家齊，2024；黃韵晴，2026）": "（周家齊，2024）",
    "（連英傑，2025；林文元，2024）": "（連英傑，2025）",
    "（陳應交，2022；廖國言，2022；連英傑，2025；郭致均，2024）": "（廖國言，2022）",
    "陳應交（2022）與簡吉岐（2024）指出": "簡吉岐（2024）指出",
    "陳應交（2022）在《違反廢棄物清理法行為主體法律責任之研究》中深入分析指出": "陳應交（2022）在《違反廢棄物清理法行為主體法律責任之研究》中深入分析指出",
    "廖國言（2022）與洪祥恩（2026）深入分析指出": "廖國言（2022）深入分析指出",
    "廖國言（2022）與洪祥恩（2026）之研究指出": "洪祥恩（2026）之研究指出",
    "陳應交（2022）、廖國言（2022）、張彩霞（2024）、連英傑（2025）、郭致均（2024）等": "連英傑（2025）、張彩霞（2024）等",
    "前述廖國言（2022）、陳應交（2022）及洪祥恩（2026）之研究所述": "前述廖國言（2022）之研究所述",
    "（連英傑，2025；張彩霞，2024）": "（連英傑，2025）",
    "張彩霞（2024）、周家齊（2024）與黃韵晴（2026）": "周家齊（2024）",
}

for old, new in replacements.items():
    content = content.replace(old, new)

# 3. Replace reference list
ref_start_marker = r"\noindent \textbf{一、中文文獻（依作者姓氏筆畫順序排列，符合 APA 第七版規範）}"
ref_end_marker = r"\noindent \textbf{二、英文文獻（Alphabetical Order by Author's Last Name, APA 7th Edition）}"
doc_end_marker = r"\section*{附錄：變數定義與資料來源對照表}"

new_chinese_refs = r"""\noindent \textbf{一、中文文獻（依作者姓氏筆畫順序排列，符合 APA 第七版規範）}

\begin{enumerate}[label={[\arabic*]}, leftmargin=3em]
    \item 李依潔（2024）。《環境構面永續發展指標、企業策略與成本僵固性》（未出版之碩士論文）。中原大學會計學系，桃園市。
    \item 林有志、傅鐘仁、陳筱平（2011）。管理當局樂觀預期、產業特質與成本僵固性。《管理學報》，28(6)，585--606。
    \item 周家齊（2024）。《環境犯罪之個人與企業刑事責任—以廢棄物清理法為例》（未出版之碩士論文）。國立臺北大學法律學系，新北市。
    \item 洪祥恩（2026）。《廢棄物清除處理機構營運管理之研究－以高雄市一家G 公司為例》（未出版之碩士論文）。嘉南藥理大學環境工程與科學系，臺南市。
    \item 張岑華（2026）。《ESG績效與漂綠對權益資金成本之影響：以台灣為例》（未出版之碩士論文）。世新大學財務金融學系，臺北市。
    \item 張彩霞（2024）。《泰國與臺灣工業廢棄物清理法制之比較研究》（未出版之碩士論文）。中央警察大學法律學研究所，桃園市。
    \item 許嘉芸（2026）。《企業資源密集度管理對財務績效與ESG效率的影響》（未出版之碩士論文）。國立陽明交通大學經營管理研究所，新竹市。
    \item 郭致均（2024）。《論煉鋼爐渣之處理與法律定性之研究-台灣與日本法制之比較》（未出版之碩士論文）。國立高雄科技大學科技法律研究所，高雄市。
    \item 連英傑（2025）。《不法傾倒事業廢棄物環境犯罪之執法困境》（未出版之碩士論文）。國立臺北大學犯罪學研究所，新北市。
    \item 陳應交（2022）。《違反廢棄物清理法行為主體法律責任之研究》（未出版之碩士論文）。嶺東科技大學財經法律研究所，臺中市。
    \item 楊博勝（2026）。《ESG資訊之揭露對於企業特有風險之影響：以台灣上市櫃公司為例》（未出版之博士論文）。國立陽明交通大學經營管理研究所，新竹市。
    \item 廖國言（2022）。《生物醫療廢棄物處理業者勞務採購契約法律爭議問題之研究－以雲林縣甲級生物醫療廢棄物處理公司為中心》（未出版之碩士論文）。國立雲林科技大學科技法律研究所，雲林縣。
    \item 簡吉岐（2024）。《土地廢棄物之清理責任——以行為責任及狀態責任為中心》（未出版之碩士論文）。國立中正大學法律學系，嘉義縣。
    \item 魏郁芳（2025）。《企業內部因素與外部壓力對水資源揭露之關聯性研究：以臺灣上市櫃公司為例》（未出版之碩士論文）。國立臺灣師範大學永續管理與環境教育研究所，臺北市。
\end{enumerate}

\vspace{1em}
"""

new_english_refs = r"""\noindent \textbf{二、英文文獻（Alphabetical Order by Author's Last Name, APA 7th Edition）}

\begin{enumerate}[label={[\arabic*]}, leftmargin=3em, resume]
    \item Anderson, M. C., Banker, R. D., \& Janakiraman, S. N. (2003). Are selling, general, and administrative costs "sticky"? \textit{Journal of Accounting Research}, 41(1), 47--63. https://doi.org/10.1111/1475-679X.00095
    \item Banker, R. D., Byzalov, D., Ciftci, M., \& Mashruwala, R. (2014). The moderating effect of prior sales changes on cost behavior. \textit{Journal of Management Accounting Research}, 26(2), 221--242. https://doi.org/10.2308/jmar-50846
    \item Chen, C. X., Lu, H., \& Sougiannis, T. (2012). The agency problem, corporate governance, and the asymmetry of cost behavior. \textit{The Accounting Review}, 87(1), 239--272. https://doi.org/10.2308/accr-10165
    \item Jiang, W., \& Yang, W. (2024). ESG disclosure and corporate cost stickiness: Evidence from supply-chain relationships. \textit{Economics Letters}, 238, Article 111697. https://doi.org/10.1016/j.econlet.2024.111697
    \item Li, P., Yang, J., Islam, M. A., \& Ren, S. (2023). Making AI less "thirsty": Uncovering and addressing the secret water footprint of AI models. \textit{arXiv preprint arXiv:2304.03271}. https://arxiv.org/abs/2304.03271
    \item Zeng, J., Peng, M., \& Chan, K. C. (2024). \textit{The impact of low-carbon city policy on corporate cost stickiness} (Working Paper). Xiangtan University \& Shanghai Business School.
\end{enumerate}

\newpage

"""

if ref_start_marker in content and doc_end_marker in content:
    before_refs = content.split(ref_start_marker)[0]
    after_refs = content.split(doc_end_marker)[1]
    content = before_refs + new_chinese_refs + new_english_refs + doc_end_marker + after_refs
else:
    print("Could not find reference markers.")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Updated 09_generate_latex.py successfully.")
