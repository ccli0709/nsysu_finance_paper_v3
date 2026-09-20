import os

filepath = "src/03_paper_generation/07_generate_paper.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_func = """def build_descriptive(df):
    rows = []
    for v in DESC_VARS:
        s = pd.to_numeric(df[v], errors="coerce").dropna()
        rows.append({"變數": v, "N": int(s.size), "平均數": s.mean(),
                     "標準差": s.std(), "最小值": s.min(),
                     "中位數": s.median(), "最大值": s.max()})
    return pd.DataFrame(rows)"""

new_func = """def build_descriptive(df):
    rows = []
    for v in DESC_VARS:
        s = pd.to_numeric(df[v], errors="coerce").dropna()
        rows.append({
            "變數": v, 
            "樣本數": int(s.size), 
            "最小值": s.min(),
            "第一四分位": s.quantile(0.25),
            "平均數": s.mean(),
            "中位數": s.median(), 
            "第三四分位": s.quantile(0.75),
            "最大值": s.max(),
            "標準差": s.std()
        })
    return pd.DataFrame(rows)"""

content = content.replace(old_func, new_func)

# Also rename role_rows in 07_generate_paper.py to match variables
# role_rows = [("β1", "?LNREV", "dLNREV"), ("β2", "D??LNREV", "DEC_x_dREV"), ...]
# I'll just regex replace "D??LNREV" with "DEC??LNREV" (Wait, ?? was × in Big5 probably, let's just replace the raw unicode or write a script to rewrite it cleanly)
