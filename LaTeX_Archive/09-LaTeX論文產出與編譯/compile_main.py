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

DIR_PATH = os.path.dirname(os.path.abspath(__file__))
if DIR_PATH not in sys.path:
    sys.path.insert(0, DIR_PATH)
import preview_server

TEX_FILE = os.path.join(DIR_PATH, "main.tex")
HTML_FILE = os.path.join(DIR_PATH, "preview_output.html")
PDF_FILE = os.path.join(DIR_PATH, "main.pdf")
import time

DOCX_FILE = os.path.join(DIR_PATH, "main.docx")

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
    
    # 檢查目標 main.pdf 是否被外部 PDF 閱讀器鎖定
    target_locked = False
    if os.path.exists(pdf_abs):
        try:
            with open(pdf_abs, 'r+b') as f:
                pass
            os.remove(pdf_abs)
        except Exception:
            target_locked = True

    # 先輸出到獨立暫存檔 main_latest.pdf，保證不受檔案鎖定干擾
    latest_pdf = os.path.join(DIR_PATH, "main_latest.pdf")
    if os.path.exists(latest_pdf):
        try:
            os.remove(latest_pdf)
        except Exception:
            pass

    cmd = [
        browser_exe,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={latest_pdf}",
        url
    ]
    
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60)
    except Exception as e:
        print(f"  [ERROR] 調用瀏覽器異常: {e}")
        return False
    
    if os.path.exists(latest_pdf) and os.path.getsize(latest_pdf) > 0:
        size_mb = os.path.getsize(latest_pdf) / (1024 * 1024)
        if not target_locked:
            try:
                import shutil
                shutil.copy2(latest_pdf, pdf_abs)
                print(f"  [SUCCESS] PDF 生成成功: {pdf_abs} ({size_mb:.2f} MB)")
                print(f"  [SUCCESS] 同步備份最新 PDF: {latest_pdf}")
                return True
            except Exception as e:
                target_locked = True
        
        if target_locked:
            print(f"  [WARNING] 偵測到 {os.path.basename(pdf_abs)} 目前正被外部 PDF 閱讀器（如 PDF-XChange Editor）開啟並鎖定中，無法直接覆寫！")
            print(f"  [SUCCESS] 最新已編譯之完整 PDF 已另存為: {latest_pdf} ({size_mb:.2f} MB)")
            print(f"  [NOTE] 請關閉 PDF 閱讀器中的 {os.path.basename(pdf_abs)} 標籤頁以允許覆寫，或直接開啟 main_latest.pdf 檢視最新論文成果。")
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
