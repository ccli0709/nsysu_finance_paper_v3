# -*- coding: utf-8 -*-
"""
compile_main.py
---------------
將 main.tex 自動編譯為 PDF (main.pdf) 與 Word (main.docx)。

流程：
1. 呼叫 preview_server 解析 main.tex 並更新 preview_output.html。
2. 使用 Microsoft Edge / Chrome 無頭瀏覽器將 HTML 渲染為高畫質 A4 PDF (main.pdf)。
3. 使用 MS Word COM 組件 (Word.Application) 將 HTML 轉存為標準 Word 檔 (main.docx)。
"""

import os
import sys
import subprocess
import preview_server

TEX_FILE = "main.tex"
HTML_FILE = "preview_output.html"
PDF_FILE = "main.pdf"
DOCX_FILE = "main.docx"

def compile_pdf(html_abs=None, pdf_abs=None):
    if not html_abs:
        html_abs = os.path.abspath(HTML_FILE)
    if not pdf_abs:
        pdf_abs = os.path.abspath(PDF_FILE)
        
    print("1/2 正在透過 Edge / Chrome 無頭瀏覽器編譯 PDF (main.pdf)...")
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
        print("錯誤：找不到 Microsoft Edge 或 Chrome 執行檔。無法生成 PDF。")
        return False
        
    url = f"file:///{html_abs.replace('\\', '/')}"
    ps_command = f'Start-Process -FilePath "{browser_exe}" -ArgumentList "--headless --print-to-pdf=\\"{pdf_abs}\\" \\"{url}\\"" -Wait'
    
    res = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
    
    if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
        size_mb = os.path.getsize(pdf_abs) / (1024 * 1024)
        print(f"  [SUCCESS] PDF 生成成功: {pdf_abs} ({size_mb:.2f} MB)")
        return True
    else:
        print(f"  [ERROR] PDF 編譯失敗。錯誤訊息：{res.stderr}")
        return False

def compile_word(html_abs=None, docx_abs=None):
    if not html_abs:
        html_abs = os.path.abspath(HTML_FILE)
    if not docx_abs:
        docx_abs = os.path.abspath(DOCX_FILE)
        
    print("2/2 正在透過 MS Word COM 組件轉換生成 Word 檔 (main.docx)...")
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 # Disable alerts/dialogs
        doc = word.Documents.Open(html_abs)
        doc.SaveAs2(docx_abs, FileFormat=16) # 16 = wdFormatDocumentDefault (.docx)
        doc.Close(False)
        word.Quit()
        
        if os.path.exists(docx_abs) and os.path.getsize(docx_abs) > 0:
            size_kb = os.path.getsize(docx_abs) / 1024
            print(f"  [SUCCESS] Word 檔生成成功: {docx_abs} ({size_kb:.2f} KB)")
            return True
        else:
            print("  [ERROR] Word 檔生成失敗，檔案大小為 0。")
            return False
    except Exception as e:
        print(f"  [ERROR] MS Word COM 轉檔失敗: {e}")
        return False

def compile_all():
    print("\n========================================================")
    print("  開始編譯 main.tex -> PDF (main.pdf) & Word (main.docx)")
    print("========================================================")
    
    if not os.path.exists(TEX_FILE):
        print(f"錯誤：找不到 {TEX_FILE}")
        return False

    print("0. 正在更新 HTML 預覽檔 (preview_output.html)...")
    if not preview_server.build_preview_file():
        print("錯誤：無法建立預覽網頁檔。")
        return False

    html_abs = os.path.abspath(HTML_FILE)
    pdf_abs = os.path.abspath(PDF_FILE)
    docx_abs = os.path.abspath(DOCX_FILE)

    pdf_ok = compile_pdf(html_abs, pdf_abs)
    word_ok = compile_word(html_abs, docx_abs)

    print("========================================================")
    if pdf_ok and word_ok:
        print("  [完成] main.tex 已成功編譯為 PDF 及 Word 文件！")
    else:
        print("  [警告] 編譯過程部分失敗，請檢查上方日誌。")
    print("========================================================\n")
    return pdf_ok and word_ok

if __name__ == "__main__":
    compile_all()
