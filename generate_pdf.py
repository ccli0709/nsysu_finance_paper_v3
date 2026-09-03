# -*- coding: utf-8 -*-
"""
generate_pdf.py
---------------
將 main.tex 解析並透過 Edge / Chrome 本地無頭瀏覽器渲染輸出為高解析度 A4 PDF 檔案 (main.pdf)。
包含完整的中文字型 (標楷體)、英文字型 (Times New Roman)、MathJax 向量數學公式與學術三線表。
"""

import os
import sys
import subprocess
import preview_server

TEX_FILE = "main.tex"
HTML_FILE = "preview_output.html"
PDF_FILE = "main.pdf"

def generate_pdf():
    print("1. 正在更新 HTML 預覽檔...")
    if not preview_server.build_preview_file():
        print("錯誤：無法建立預覽網頁檔。")
        return False
        
    html_abs = os.path.abspath(HTML_FILE)
    pdf_abs = os.path.abspath(PDF_FILE)
    
    print("2. 正在透過 Microsoft Edge 無頭瀏覽器編譯 PDF...")
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    ]
    
    browser_exe = None
    for path in edge_paths:
        if os.path.exists(path):
            browser_exe = path
            break
            
    if not browser_exe:
        print("錯誤：找不到 Microsoft Edge 或 Chrome 執行檔。")
        return False
        
    url = f"file:///{html_abs.replace('\\', '/')}"
    
    ps_command = f'Start-Process -FilePath "{browser_exe}" -ArgumentList "--headless --print-to-pdf=\\"{pdf_abs}\\" \\"{url}\\"" -Wait'
    
    res = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
    
    if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
        size_mb = os.path.getsize(pdf_abs) / (1024 * 1024)
        print(f"\n========================================================")
        print(f"  成功生成高畫質 PDF 完稿文件！")
        print(f"  檔案路徑: {pdf_abs}")
        print(f"  檔案大小: {size_mb:.2f} MB")
        print(f"========================================================\n")
        return True
    else:
        print(f"PDF 編譯失敗。錯誤訊息：{res.stderr}")
        return False

if __name__ == "__main__":
    generate_pdf()
