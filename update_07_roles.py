import os

filepath = "src/03_paper_generation/07_generate_paper.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('"D×ΔLNREV"', '"DEC×ΔLNREV"')
content = content.replace('"??變數×D×ΔLNREV"', '"??變數×DEC×ΔLNREV"')
content = content.replace('"SIZE×D×ΔLNREV"', '"SIZE×DEC×ΔLNREV"')
content = content.replace('"AI×D×ΔLNREV"', '"AI×DEC×ΔLNREV"')
content = content.replace('"EI×D×ΔLNREV"', '"EI×DEC×ΔLNREV"')
content = content.replace('"ROA×D×ΔLNREV"', '"ROA×DEC×ΔLNREV"')
content = content.replace('"LEV×D×ΔLNREV"', '"LEV×DEC×ΔLNREV"')
content = content.replace('"Decrease×D×ΔLNREV"', '"SUCC_DEC×DEC×ΔLNREV"')
content = content.replace('"SUCC_DEC×D×ΔLNREV"', '"SUCC_DEC×DEC×ΔLNREV"')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
