# -*- coding: utf-8 -*-
"""
update_thesis.py
----------------
以「論文生成結果/編輯版本.docx」為母本（保留作者親自修訂之第 1、2、3、5 章及所有版面排版），
將第四章實證結果全文與 5 個實證表格精準替換為：
1. 表格變數純英數字代號（Alphanumeric Variable Codes），無任何中文變數名稱。
2. 內文風格完全對齊修訂後的第 1、2、3、5 章，客觀專業、符合台灣會計與財務計量論文標準。
3. 變數全面 Word 原生行內公式 (<m:oMath>)。
4. 標題層級綁定 Heading 1 / 2 / 3，方便大綱導覽。
5. 全文禁忌詞彙校正（校正 P283 遺漏之「嚴謹」為「客觀標準」，修復雙句號）。
產出最新正式版論文 Word 檔。
"""

import os
import copy
import docx
import pandas as pd
from docx import Document
from generate_full_paper import (
    MODEL_CSV, COEF_CSV, ROB_CSV,
    build_chapter_4
)

EDITED_DOC_PATH = r"d:\Repositories\nsysu_finance_paper_v3\論文生成結果\編輯版本.docx"
OUTPUT_DOC_PATH = r"d:\Repositories\nsysu_finance_paper_v3\論文生成結果\企業水資源管理與廢棄物處理對成本僵固性之影響_20261114_正式修訂版.docx"

def main():
    print("1. 讀取實證數據...")
    df = pd.read_csv(MODEL_CSV, low_memory=False)
    coef_df = pd.read_csv(COEF_CSV)
    rob_df = pd.read_csv(ROB_CSV)
    print("   數據讀取完成。")

    print("2. 讀取母本：編輯版本.docx...")
    doc = Document(EDITED_DOC_PATH)
    body = doc.element.body

    # 定位第四章與第五章起點
    ch4_start = None
    ch5_start = None
    for idx, child in enumerate(body):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            p_text = child.xpath('string(.)').strip()
            if p_text == '第四章 實證結果與分析':
                ch4_start = idx
            elif p_text == '第五章 結論與建議':
                ch5_start = idx
                break

    print(f"   定位成功：第四章起始 child 索引 = {ch4_start}，第五章起始 child 索引 = {ch5_start}")
    assert ch4_start is not None and ch5_start is not None, "未找到第四章或第五章起點！"

    print("3. 在臨時文檔中構建最新第四章（純英數字表格 + 專業計量內文 + OMML）...")
    temp_doc = Document()
    build_chapter_4(temp_doc, df, coef_df, rob_df)

    # 設置標題樣式 (Heading 1, 2, 3) 確保導覽窗格與目錄可正常解析
    for p in temp_doc.paragraphs:
        txt = p.text.strip()
        if txt == '第四章 實證結果與分析':
            p.style = doc.styles['Heading 1']
        elif any(txt.startswith(h) for h in ['第一節', '第二節', '第三節', '第四節', '第五節', '第六節']):
            p.style = doc.styles['Heading 2']
        elif any(txt.startswith(h) for h in ['一、', '二、', '三、', '（一）', '（二）', '（三）']):
            # 判斷是否為子節標題（文字短且粗體）
            if len(txt) < 50 and ('加劇' in txt or '對照' in txt):
                p.style = doc.styles['Heading 3']

    new_ch4_elements = list(temp_doc.element.body)
    # 移除 temp_doc 結尾的 sectPr（若有）
    if new_ch4_elements and new_ch4_elements[-1].tag.endswith('sectPr'):
        new_ch4_elements.pop()

    print(f"   最新第四章生成完成，共 {len(new_ch4_elements)} 個 XML 節點（段落與表格）。")

    print("4. 替換舊版第四章節點...")
    old_count = ch5_start - ch4_start
    for _ in range(old_count):
        body.remove(body[ch4_start])
    print(f"   已刪除舊第四章 {old_count} 個 XML 節點。")

    ch5_node = body[ch4_start]
    for elem in new_ch4_elements:
        elem_clone = copy.deepcopy(elem)
        ch5_node.addprevious(elem_clone)
    print(f"   已成功插入新第四章 {len(new_ch4_elements)} 個 XML 節點！")

    print("5. 檢查並校正第 5 章殘存之禁用詞彙與格式細節...")
    for p in doc.paragraphs:
        if '研訂更為嚴謹之' in p.text:
            for r in p.runs:
                if '研訂更為嚴謹之' in r.text:
                    r.text = r.text.replace('研訂更為嚴謹之', '研訂更為客觀標準之')
                if '漂綠風險。。' in r.text:
                    r.text = r.text.replace('漂綠風險。。', '漂綠風險。')
            print("   已校正 P283 政策建議中之『嚴謹』為『客觀標準』並修正連續雙句號。")

    print(f"6. 儲存新版論文至：{OUTPUT_DOC_PATH}...")
    doc.save(OUTPUT_DOC_PATH)
    print("   新版論文儲存成功！")

if __name__ == "__main__":
    main()
