import os
import glob
import re

filepath = "src/03_paper_generation/09_generate_latex.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace $\Delta\text{LNSGA}$ with $\Delta LNSGA$
content = content.replace(r'$\Delta\text{LNSGA}$', r'$\Delta LNSGA$')
content = content.replace(r'$\Delta\text{LNREV}$', r'$\Delta LNREV$')
content = content.replace(r'$D$', r'$DEC$')
content = content.replace(r'\text{Water\_Intensity}', r'WATER\_INTENSITY')
content = content.replace(r'\text{Waste\_Intensity}', r'WASTE\_INTENSITY')
content = content.replace(r'\text{Water\_Rate}', r'WATER\_RATE')
content = content.replace(r'\text{Waste\_Disc}', r'WASTE\_DISC')
content = content.replace(r'\text{Waste\_Fine}', r'WASTE\_FINE')
content = content.replace(r'\text{SIZE}', r'SIZE')
content = content.replace(r'\text{AI}', r'AI')
content = content.replace(r'\text{EI}', r'EI')
content = content.replace(r'\text{ROA}', r'ROA')
content = content.replace(r'\text{LEV}', r'LEV')
content = content.replace(r'\text{SUCC_DEC}', r'SUCC\_DEC')
content = content.replace(r'\text{Water\_Int}', r'WATER\_INTENSITY')
content = content.replace(r'\text{Waste\_Int}', r'WASTE\_INTENSITY')

# Table format for descriptive statistics
# Currently it's likely generating N, Mean, P25, Median, P75, Max.
# The reference table is: 樣本數, 最小值, 第一四分位, 平均數, 中位數, 第三四分位, 最大值, 標準差
# We will just write a wrapper around it to change the column headers.

content = content.replace(r'變數 & 樣本數 & 平均數 & 標準差 & 最小值 & P25 & 中位數 & P75 & 最大值 \\', 
                          r'變數 & 樣本數 & 最小值 & 第一四分位 & 平均數 & 中位數 & 第三四分位 & 最大值 & 標準差 \\')

content = content.replace(r'v_mean = df[col].mean()', r'v_mean = df[col].mean()')
content = content.replace(r'v_std = df[col].std()', r'v_std = df[col].std()')
# Let's adjust the row format string if we can find it
# It's probably formatted like: {v_mean:.3f} & {v_std:.3f} ...

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

