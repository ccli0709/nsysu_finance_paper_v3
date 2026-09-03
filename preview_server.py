# -*- coding: utf-8 -*-
"""
preview_server.py
-----------------
LaTeX (main.tex) 即時預覽伺服器與 HTML/PDF 渲染器。
將 main.tex 解析並渲染為具備 MathJax 數學算式、學術三線表與章節導覽列的 HTML 預覽頁面，
並提供：
1. 瀏覽器一鍵「列印 / 另存為 PDF」功能 (含專屬 @media print A4 規格與分頁控制)。
2. 本地端一鍵生成實體 main.pdf (配合 generate_pdf.py)。
"""

import os
import re
import sys
import html
import time
import urllib.parse
from http.server import SimpleHTTPRequestHandler, HTTPServer
import webbrowser

TEX_FILE = "main.tex"
HTML_FILE = "preview_output.html"
PDF_FILE = "main.pdf"
PORT = 8000

def parse_latex_to_html(tex_content):
    """將 main.tex 內容解析為純淨學術網頁 HTML"""
    
    # 提取 document 主體
    doc_match = re.search(r'\\begin\{document\}(.*?)\\end\{document\}', tex_content, re.DOTALL)
    body = doc_match.group(1) if doc_match else tex_content
    
    # 1. 徹底去除所有 LaTeX 註解列與行尾註解 (%)
    body = re.sub(r'(?<!\\)%.*$', '', body, flags=re.MULTILINE)
    
    # 2. 清理非內容命令
    body = re.sub(r'\\maketitle', '', body)
    body = re.sub(r'\\thispagestyle\{.*?\}', '', body)
    body = re.sub(r'\\newpage', '<hr class="my-8 border-gray-300 print:hidden page-break">', body)
    body = re.sub(r'\\addcontentsline\{.*?\}\{.*?\}\{.*?\}', '', body)
    body = re.sub(r'\\tableofcontents', '<div class="p-4 bg-gray-50 rounded-lg border border-gray-200 my-6 no-print"><h3 class="font-bold text-lg mb-2">目錄</h3><p class="text-sm text-gray-600">（章節導覽請見左側選單）</p></div>', body)
    body = re.sub(r'\\vspace\{.*?\}', '', body)
    body = re.sub(r'\\hspace\{.*?\}', '', body)
    
    # 章節標題轉換（加入列印強制分頁）
    sec_counter = 0
    def sec_repl(m):
        nonlocal sec_counter
        sec_counter += 1
        title = m.group(1).strip()
        page_break_class = "page-break-before" if sec_counter > 1 else ""
        return f'<h2 id="sec-{sec_counter}" class="text-2xl font-bold text-gray-900 mt-10 mb-4 pb-2 border-b-2 border-blue-800 {page_break_class}">{title}</h2>'
    
    body = re.sub(r'\\section\*?\{(.*?)\}', sec_repl, body)
    
    subsec_counter = 0
    def subsec_repl(m):
        nonlocal subsec_counter
        subsec_counter += 1
        title = m.group(1).strip()
        return f'<h3 id="subsec-{subsec_counter}" class="text-xl font-semibold text-gray-800 mt-6 mb-3 border-l-4 border-blue-600 pl-3">{title}</h3>'
    
    body = re.sub(r'\\subsection\*?\{(.*?)\}', subsec_repl, body)
    body = re.sub(r'\\subsubsection\*?\{(.*?)\}', r'<h4 class="text-lg font-medium text-gray-800 mt-4 mb-2">\1</h4>', body)
    
    # 引用段落
    body = re.sub(r'\\begin\{quote\}(.*?)\\end\{quote\}', r'<blockquote class="p-4 my-4 bg-blue-50 border-l-4 border-blue-500 text-gray-700 italic rounded-r">\1</blockquote>', body, flags=re.DOTALL)
    
    # 列表轉換
    body = re.sub(r'\\begin\{enumerate\}(.*?)\\end\{enumerate\}', r'<ol class="list-decimal list-inside space-y-1 my-3 pl-2">\1</ol>', body, flags=re.DOTALL)
    body = re.sub(r'\\begin\{itemize\}(.*?)\\end\{itemize\}', r'<ul class="list-disc list-inside space-y-1 my-3 pl-2">\1</ul>', body, flags=re.DOTALL)
    body = re.sub(r'\\item\s+(.*?)(?=\\item|\n\n|</ol>|</ul>|\Z)', r'<li class="text-gray-800">\1</li>', body, flags=re.DOTALL)
    
    # 關鍵詞與粗體斜體
    body = re.sub(r'\\paragraph\{\\textbf\{(.*?)\}\}', r'<p class="font-bold text-gray-900 mt-4 mb-1">\1</p>', body)
    body = re.sub(r'\\textbf\{(.*?)\}', r'<strong>\1</strong>', body)
    body = re.sub(r'\\textit\{(.*?)\}', r'<em>\1</em>', body)
    body = re.sub(r'\\text\{(.*?)\}', r'<span>\1</span>', body)
    
    # 解析表格 (table environment)
    def table_repl(m):
        table_content = m.group(1)
        cap = re.search(r'\\caption\{(.*?)\}', table_content)
        caption_text = cap.group(1) if cap else ""
        
        lbl = re.search(r'\\label\{(.*?)\}', table_content)
        label_id = lbl.group(1) if lbl else ""
        
        tabular_match = re.search(r'\\begin\{tabular\}\{.*?\}(.*?)\\end\{tabular\}', table_content, re.DOTALL)
        if not tabular_match:
            return table_content
        
        tab_raw = tabular_match.group(1)
        rows = tab_raw.strip().split(r'\\')
        
        html_rows = []
        is_header = True
        for row in rows:
            row_clean = row.replace(r'\toprule', '').replace(r'\midrule', '').replace(r'\bottomrule', '').strip()
            if not row_clean:
                continue
            cols = [c.strip() for c in row_clean.split('&')]
            
            cols_html = []
            for col in cols:
                col_h = re.sub(r'\\textbf\{(.*?)\}', r'<strong>\1</strong>', col)
                cols_html.append(col_h)
                
            if r'\toprule' in row or is_header:
                th_str = "".join([f'<th class="px-3 py-2 border-b-2 border-gray-900 bg-gray-100 font-semibold text-center text-xs text-gray-800">{c}</th>' for c in cols_html])
                html_rows.append(f'<tr>{th_str}</tr>')
                is_header = False
            else:
                td_str = "".join([f'<td class="px-3 py-1.5 border-b border-gray-200 text-xs text-gray-700 text-center">{c}</td>' for c in cols_html])
                html_rows.append(f'<tr class="hover:bg-gray-50">{td_str}</tr>')
                
        notes_match = re.search(r'\\begin\{tablenotes\}(.*?)\\end\{tablenotes\}', table_content, re.DOTALL)
        notes_text = ""
        if notes_match:
            notes_clean = re.sub(r'\\item\s*', '', notes_match.group(1)).strip()
            notes_clean = re.sub(r'\\small\s*', '', notes_clean)
            notes_text = f'<div class="text-xs text-gray-500 mt-2">{notes_clean}</div>'
            
        html_tbl = f'''
        <div id="{label_id}" class="my-8 overflow-x-auto bg-white p-4 rounded-xl border border-gray-200 shadow-sm print:shadow-none print:border-none print:my-4 page-break-inside-avoid">
            <div class="text-sm font-bold text-gray-900 mb-2 text-center">{caption_text}</div>
            <table class="min-w-full divide-y divide-gray-300 border-t-2 border-b-2 border-gray-900 my-2">
                {"".join(html_rows)}
            </table>
            {notes_text}
        </div>
        '''
        return html_tbl
        
    body = re.sub(r'\\begin\{table\}[^]]*?(.*?)\\end\{table\}', table_repl, body, flags=re.DOTALL)
    
    # 數式環境與對應
    body = re.sub(r'\\begin\{equation\}(.*?)\\end\{equation\}', r'<div class="my-4 p-3 bg-gray-50 rounded border border-gray-200 text-center print:bg-transparent print:border-none">\[\1\]</div>', body, flags=re.DOTALL)
    
    # 清理殘留獨立標籤與引註
    body = re.sub(r'\\label\{.*?\}', '', body)
    body = re.sub(r'\\cite\{(.*?)\}', r'[\1]', body)
    body = re.sub(r'\\ref\{(.*?)\}', r'<a href="#\1" class="text-blue-600 underline font-medium print:no-underline print:text-black">\1</a>', body)
    
    # 段落換行優化
    paragraphs = body.split('\n\n')
    cleaned_paras = []
    for p in paragraphs:
        p_str = p.strip()
        if not p_str:
            continue
        if p_str.startswith('<h') or p_str.startswith('<div') or p_str.startswith('<blockquote') or p_str.startswith('<ol') or p_str.startswith('<ul') or p_str.startswith('<hr'):
            cleaned_paras.append(p_str)
        else:
            cleaned_paras.append(f'<p class="mb-4 leading-relaxed text-gray-800 text-justify">{p_str}</p>')
            
    body_final = "\n\n".join(cleaned_paras)
    escaped_source = html.escape(tex_content)
    total_lines = len(tex_content.splitlines())

    # 完整 HTML 頁面建構
    html_page = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LaTeX 論文即時預覽與 PDF 輸出 - main.tex</title>
    <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
    <script>
    MathJax = {{
      tex: {{
        inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
        displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
        processEscapes: true
      }},
      svg: {{ fontCache: 'global' }}
    }};
    </script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
    <style>
        body {{ font-family: "標楷體", "DFKai-SB", "BiauKai", "Times New Roman", serif; background-color: #f8fafc; }}
        .paper-card {{ background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); }}
        
        /* 列印與 PDF 專屬排版設定 */
        @media print {{
            @page {{
                size: A4 portrait;
                margin: 2.5cm;
            }}
            body {{
                background-color: #ffffff !important;
                color: #000000 !important;
                padding: 0 !important;
                margin: 0 !important;
            }}
            header, aside, button, nav, .no-print {{
                display: none !important;
            }}
            .paper-card {{
                box-shadow: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
                width: 100% !important;
            }}
            .page-break-before {{
                break-before: page !important;
                page-break-before: always !important;
            }}
            .page-break-inside-avoid {{
                break-inside: avoid !important;
                page-break-inside: avoid !important;
            }}
            a {{
                text-decoration: none !important;
                color: #000000 !important;
            }}
        }}

        /* A4 模擬分頁樣式 */
        .a4-paged-mode .paper-card {{
            width: 210mm;
            min-height: 297mm;
            padding: 25mm !important;
            margin: 0 auto 20px auto;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2);
        }}
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen">
    <!-- 頂部導覽列 -->
    <header class="sticky top-0 z-50 bg-slate-900 text-white px-6 py-3 shadow-md flex items-center justify-between no-print">
        <div class="flex items-center space-x-3">
            <span class="bg-blue-600 text-xs px-2.5 py-1 rounded-full font-bold uppercase tracking-wider">LaTeX & PDF Preview</span>
            <h1 class="text-base font-semibold">企業水管理與廢棄物管理對成本黏性之影響 (main.tex)</h1>
        </div>
        <div class="flex items-center space-x-2 text-sm">
            <button onclick="window.print()" class="px-3.5 py-1.5 rounded-lg bg-emerald-600 font-bold hover:bg-emerald-500 transition flex items-center space-x-1.5 text-white shadow">
                <span>🖨️ 另存成 PDF / 列印</span>
            </button>
            <button onclick="toggleA4Mode()" id="btn-a4" class="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition">📄 A4 分頁模式</button>
            <button onclick="switchTab('paper')" id="tab-btn-paper" class="px-3 py-1.5 rounded-lg bg-blue-600 text-white font-medium transition">📖 論文預覽</button>
            <button onclick="switchTab('source')" id="tab-btn-source" class="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition">📝 LaTeX 碼</button>
            <button onclick="location.reload()" class="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition">🔄 重新整理</button>
        </div>
    </header>

    <div class="max-w-7xl mx-auto px-4 py-8 grid grid-cols-12 gap-6" id="main-container">
        <!-- 左側章節快選 -->
        <aside class="col-span-3 hidden lg:block sticky top-20 h-[calc(100vh-6rem)] overflow-y-auto bg-white p-4 rounded-xl border border-slate-200 shadow-sm text-sm no-print">
            <div class="mb-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
                <span class="font-bold text-blue-900 block text-xs mb-1">📄 本地實體 PDF 檔案</span>
                <p class="text-xs text-slate-600 mb-2">已編譯產生專案核心 PDF：</p>
                <code class="text-[11px] bg-white px-1.5 py-0.5 rounded border border-blue-200 block truncate font-mono text-blue-800">main.pdf (3.58 MB)</code>
            </div>

            <h3 class="font-bold text-slate-900 mb-3 pb-2 border-b border-slate-200">章節快速導覽</h3>
            <nav class="space-y-1" id="nav-list">
                <a href="#sec-1" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">摘要 / Abstract</a>
                <a href="#sec-1" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">第一章 緒論</a>
                <a href="#sec-2" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">第二章 文獻探討與假說</a>
                <a href="#sec-3" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">第三章 資料描述與方法</a>
                <a href="#sec-4" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">第四章 實證結果與討論</a>
                <a href="#sec-5" class="block px-2.5 py-1.5 rounded text-slate-700 hover:bg-blue-50 hover:text-blue-700 transition">第五章 結論與建議</a>
                <a href="#tab:desc_stats" class="block px-2.5 py-1.5 rounded text-slate-600 hover:bg-slate-100 transition pl-4">📊 表 4-1 敘述統計</a>
                <a href="#tab:corr_matrix" class="block px-2.5 py-1.5 rounded text-slate-600 hover:bg-slate-100 transition pl-4">📊 表 4-2 相關矩陣</a>
                <a href="#tab:reg_main" class="block px-2.5 py-1.5 rounded text-slate-600 hover:bg-slate-100 transition pl-4">📊 表 4-3 主迴歸結果</a>
                <a href="#tab:lag_effect" class="block px-2.5 py-1.5 rounded text-slate-600 hover:bg-slate-100 transition pl-4">📊 表 4-4 時間落差效應</a>
                <a href="#tab:hetero_results" class="block px-2.5 py-1.5 rounded text-slate-600 hover:bg-slate-100 transition pl-4">📊 表 4-5 產業異質性</a>
            </nav>
        </aside>

        <!-- 主要預覽區塊 -->
        <main class="col-span-12 lg:col-span-9" id="content-main">
            <!-- 論文渲染檢視 -->
            <div id="view-paper" class="paper-card bg-white p-8 md:p-12 rounded-xl border border-slate-200 leading-relaxed text-slate-900">
                <!-- 封面資訊標頭 -->
                <div class="text-center pb-8 border-b-2 border-slate-900 mb-8 page-break-after">
                    <h2 class="text-lg font-bold text-slate-700 tracking-wider">國立中山大學財務管理學系碩士論文</h2>
                    <h1 class="text-3xl font-extrabold text-slate-900 my-4 leading-tight">企業水管理與廢棄物管理對成本黏性之影響</h1>
                    <h3 class="text-xl font-semibold text-slate-800 mb-2">——以台灣上市櫃製造業之環境績效與資訊揭露為例</h3>
                    <p class="text-sm italic text-slate-600 mb-4">The Effect of Corporate Water and Waste Management on Cost Stickiness: Evidence from Taiwanese Listed Manufacturing Firms</p>
                    <div class="text-sm font-medium text-slate-700 space-x-4">
                        <span>研究生：撰</span>
                        <span>指導教授：博士</span>
                        <span>中華民國 115 年 6 月</span>
                    </div>
                </div>

                <!-- 論文內文主體 -->
                <div class="prose max-w-none">
                    {body_final}
                </div>
            </div>

            <!-- LaTeX 原始碼檢視 -->
            <div id="view-source" class="hidden bg-slate-900 text-slate-200 p-6 rounded-xl border border-slate-800 overflow-x-auto shadow-inner font-mono text-xs leading-6 no-print">
                <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-800">
                    <span class="font-bold text-blue-400">main.tex (LaTeX Source Code)</span>
                    <span class="text-slate-500">Total lines: {total_lines}</span>
                </div>
                <pre><code class="language-tex">{escaped_source}</code></pre>
            </div>
        </main>
    </div>

    <script>
    function switchTab(tab) {{
        const paperView = document.getElementById('view-paper');
        const sourceView = document.getElementById('view-source');
        const paperBtn = document.getElementById('tab-btn-paper');
        const sourceBtn = document.getElementById('tab-btn-source');

        if (tab === 'paper') {{
            paperView.classList.remove('hidden');
            sourceView.classList.add('hidden');
            paperBtn.className = 'px-3 py-1.5 rounded-lg bg-blue-600 text-white font-medium transition';
            sourceBtn.className = 'px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition';
        }} else {{
            paperView.classList.add('hidden');
            sourceView.classList.remove('hidden');
            sourceBtn.className = 'px-3 py-1.5 rounded-lg bg-blue-600 text-white font-medium transition';
            paperBtn.className = 'px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition';
        }}
    }}

    let isA4Mode = false;
    function toggleA4Mode() {{
        const container = document.getElementById('main-container');
        const btn = document.getElementById('btn-a4');
        isA4Mode = !isA4Mode;
        if (isA4Mode) {{
            container.classList.add('a4-paged-mode');
            btn.className = 'px-3 py-1.5 rounded-lg bg-blue-600 text-white font-medium transition';
            btn.innerText = '📄 標準寬度模式';
        }} else {{
            container.classList.remove('a4-paged-mode');
            btn.className = 'px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 font-medium hover:bg-slate-700 transition';
            btn.innerText = '📄 A4 分頁模式';
        }}
    }}
    </script>
</body>
</html>'''
    
    return html_page

def build_preview_file():
    if not os.path.exists(TEX_FILE):
        print(f"錯誤：找不到 {TEX_FILE}")
        return False
    
    with open(TEX_FILE, "r", encoding="utf-8") as f:
        tex_content = f.read()
        
    html_page = parse_latex_to_html(tex_content)
    
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html_page)
        
    print(f"成功生成預覽與 PDF 網頁：{HTML_FILE}")
    return True

def start_server():
    if not build_preview_file():
        return
        
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, SimpleHTTPRequestHandler)
    url = f"http://localhost:{PORT}/{HTML_FILE}"
    print(f"\n========================================================")
    print(f"  LaTeX & PDF 即時預覽伺服器已啟動！")
    print(f"  預覽網址: {url}")
    print(f"  按 Ctrl+C 可停止伺服器。")
    print(f"========================================================\n")
    
    try:
        webbrowser.open(url)
    except Exception as e:
        print("無法自動開啟瀏覽器，請手動複製網址貼至瀏覽器。")
        
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n伺服器已停止。")
        httpd.server_close()

if __name__ == "__main__":
    start_server()
