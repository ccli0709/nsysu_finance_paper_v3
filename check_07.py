import os
import re

filepath = "src/03_paper_generation/07_generate_paper.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# descriptive stats table headers
content = content.replace('["變數", "N", "平均數", "標準差", "中位數", "P25", "P75"]', '["變數", "樣本數", "最小值", "第一四分位", "平均數", "中位數", "第三四分位", "最大值", "標準差"]')
content = content.replace('["變數名稱", "N", "平均數", "標準差", "中位數", "P25", "P75"]', '["變數", "樣本數", "最小值", "第一四分位", "平均數", "中位數", "第三四分位", "最大值", "標準差"]')

# finding where values are put into the table:
# row_cells[1].text = f"{n_val:,}"
# row_cells[2].text = f"{mean_val:.3f}"
# ... we can just replace the block since it's standard python

# Let's search and replace with regex for the descriptive stats block.
