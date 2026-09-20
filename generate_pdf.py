# -*- coding: utf-8 -*-
"""
generate_pdf.py
---------------
編譯 main.tex 輸出為 main.pdf 及 main.docx。
(相容舊指令名稱 generate_pdf.py，底層呼叫 compile_main)
"""

import compile_main

def generate_pdf():
    return compile_main.compile_all()

if __name__ == "__main__":
    generate_pdf()
