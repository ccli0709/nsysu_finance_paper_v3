# -*- coding: utf-8 -*-
"""
generate_pdf.py
---------------
編譯 main.tex 輸出為 main.pdf 及 main.docx。
(相容舊指令名稱 generate_pdf.py，底層呼叫 compile_main)
"""

import os
import sys

DIR_PATH = os.path.dirname(os.path.abspath(__file__))
if DIR_PATH not in sys.path:
    sys.path.insert(0, DIR_PATH)
import compile_main

def generate_pdf():
    return compile_main.compile_all()

if __name__ == "__main__":
    generate_pdf()
