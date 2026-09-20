import os
import re

filepath = "src/03_paper_generation/07_generate_paper.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# We need to replace the column headers array for descriptive stats
content = content.replace('["變數名稱", "N", "平均數", "標準差", "中位數", "P25", "P75"]', '["變數", "樣本數", "最小值", "第一四分位", "平均數", "中位數", "第三四分位", "最大值", "標準差"]')
content = content.replace('["變數", "N", "平均數", "標準差", "中位數", "P25", "P75"]', '["變數", "樣本數", "最小值", "第一四分位", "平均數", "中位數", "第三四分位", "最大值", "標準差"]')

# And the rows logic
#         row_cells[1].text = f"{n_val:,}"
#         row_cells[2].text = f"{mean_val:.3f}"
#         row_cells[3].text = f"{std_val:.3f}"
#         row_cells[4].text = f"{med_val:.3f}"
#         row_cells[5].text = f"{p25_val:.3f}"
#         row_cells[6].text = f"{p75_val:.3f}"

# Let's just find and replace the whole block
old_block = """        s = df[col].dropna()
        n_val = len(s)
        mean_val = s.mean()
        std_val = s.std()
        med_val = s.median()
        p25_val = s.quantile(0.25)
        p75_val = s.quantile(0.75)
        
        row_cells = t.add_row().cells
        row_cells[0].text = desc_name
        row_cells[1].text = f"{n_val:,}"
        row_cells[2].text = f"{mean_val:.4f}"
        row_cells[3].text = f"{std_val:.4f}"
        row_cells[4].text = f"{med_val:.4f}"
        row_cells[5].text = f"{p25_val:.4f}"
        row_cells[6].text = f"{p75_val:.4f}"""

new_block = """        s = df[col].dropna()
        n_val = len(s)
        min_val = s.min()
        p25_val = s.quantile(0.25)
        mean_val = s.mean()
        med_val = s.median()
        p75_val = s.quantile(0.75)
        max_val = s.max()
        std_val = s.std()
        
        row_cells = t.add_row().cells
        row_cells[0].text = desc_name
        row_cells[1].text = f"{n_val:,}"
        row_cells[2].text = f"{min_val:.3f}"
        row_cells[3].text = f"{p25_val:.3f}"
        row_cells[4].text = f"{mean_val:.3f}"
        row_cells[5].text = f"{med_val:.3f}"
        row_cells[6].text = f"{p75_val:.3f}"
        row_cells[7].text = f"{max_val:.3f}"
        row_cells[8].text = f"{std_val:.3f}"""

content = content.replace(old_block, new_block)

# Fix ROLE_ROWS if it exists in 07
content = content.replace(r'D \times \Delta LNREV', r'DEC_{i,t} \times \Delta LNREV_{i,t}')
# And others as needed, but in python they might not be latex formulas, they might be word text
# Let's just write this file
with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
