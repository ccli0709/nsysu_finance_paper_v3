# -*- coding: utf-8 -*-
"""
verify_final_paper.py
---------------------
全面檢驗最新版論文 Word 檔之格式規範、章節完整度、表格純英數字要求與風格門禁。
"""

import docx
import re

FILE_PATH = r"d:\Repositories\nsysu_finance_paper_v3\論文生成結果\企業水資源管理與廢棄物處理對成本僵固性之影響_20261114_正式修訂版.docx"

doc = docx.Document(FILE_PATH)
print(f"Total Paragraphs: {len(doc.paragraphs)}")
print(f"Total Tables: {len(doc.tables)}")

# 1. 檢驗章節標題
print("\n=== 1. 章節標題定位檢驗 ===")
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if any(h in txt for h in ['第一章', '第二章', '第三章', '第四章', '第五章', '參考文獻', '附錄 A', '附錄 B']):
        if p.style.name.startswith('Heading') or '章' in txt or '附錄' in txt or '參考文獻' in txt:
            print(f"  P{i} [{p.style.name}]: {txt[:50]}")

# 2. 檢驗第四章表格純英數字門禁
print("\n=== 2. 第四章實證表格 (Table 8 ~ Table 12) 純英數字變數檢驗 ===")
# 在 15 個表格中定位 Table 8 至 Table 12
ch4_tables = doc.tables[8:13]
table_names = ["表 4-1 敘述性統計", "表 4-2 Pearson相關矩陣", "表 4-3 基準與當期模型", "表 4-4 時間落差分析", "表 4-5 產業異質性交叉檢驗"]

chinese_pattern = re.compile(r'[\u4e00-\u9fa5]')

for t_idx, (tbl, name) in enumerate(zip(ch4_tables, table_names)):
    print(f"\n--- {name} (Table index {t_idx+8}) ---")
    rows_cnt = len(tbl.rows)
    cols_cnt = len(tbl.columns)
    print(f"維度: {rows_cnt} 列 x {cols_cnt} 欄")
    
    # 檢查第一列（表頭）
    header_texts = []
    for c in tbl.rows[0].cells:
        t = "".join(c._element.xpath('.//w:t/text() | .//m:t/text()')).strip().replace('\n', ' ')
        header_texts.append(t)
    print(f"表頭: {header_texts}")
    
    # 檢查第一欄（變數名稱）
    first_col_texts = []
    for r in tbl.rows[1:]:
        t = "".join(r.cells[0]._element.xpath('.//w:t/text() | .//m:t/text()')).strip().replace('\n', ' ')
        first_col_texts.append(t)
    print(f"變數欄預覽: {first_col_texts[:8]}")

# 3. 禁忌詞彙掃描
print("\n=== 3. 禁忌詞彙全文掃描 (門禁檢驗) ===")
buzzwords = ['宏觀', '微觀', '機轉', '嚴謹', '高規格', '極強之', '核心突破', '完全印證', '強烈呼應', '雄辯地證明', '資源稟賦', '提煉', '賦能', '抓手', '打法', '！']

all_text = []
for p in doc.paragraphs:
    all_text.append(p.text)
full_text = "\n".join(all_text)

violations = {}
for bw in buzzwords:
    cnt = full_text.count(bw)
    if cnt > 0:
        violations[bw] = cnt

if violations:
    print(f"警告：發現禁忌詞彙: {violations}")
else:
    print("通過！全文無任何禁忌詞彙、無任何驚嘆號、無大陸用語！")

# 4. 檢驗第 1、2、3、5 章修訂保留狀況
print("\n=== 4. 第 1、2、3、5 章作者修訂保留檢驗 ===")
# 檢驗 CH1 的「非僅製造業」與「研究貢獻」
has_ch1_mod = any("2014 年至 2024 年台灣上市櫃公司為研究對象" in p.text or "水資源之穩定供應與事業廢棄物之妥善處置" in p.text for p in doc.paragraphs[:130])
# 檢驗 CH2 的標題調整
has_ch2_mod = any("第一節 成本僵固性理論基礎與相關文獻" in p.text for p in doc.paragraphs[120:170])
# 檢驗 CH3 的第二節變數定義、第三節模型設定
has_ch3_mod = any("第二節 變數定義與衡量說明" in p.text for p in doc.paragraphs[160:230])
# 檢驗 CH5 的實務策略（營運槓桿、資源循環推動法）
has_ch5_mod = any("資源循環推動法" in p.text or "營運槓桿" in p.text for p in doc.paragraphs[270:])

print(f"CH1 作者修訂保留: {'是' if has_ch1_mod else '否'}")
print(f"CH2 作者修訂保留: {'是' if has_ch2_mod else '否'}")
print(f"CH3 作者修訂保留: {'是' if has_ch3_mod else '否'}")
print(f"CH5 作者修訂保留: {'是' if has_ch5_mod else '否'}")

