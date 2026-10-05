import re
import sys

def latex_to_md(tex):
    # Remove preamble and document environment
    match = re.search(r'\\begin\{document\}(.*?)\\end\{document\}', tex, re.DOTALL)
    if match:
        tex = match.group(1)
        
    # Headers
    tex = re.sub(r'\\section\*?\{(.*?)\}', r'# \1', tex)
    tex = re.sub(r'\\subsection\*?\{(.*?)\}', r'## \1', tex)
    tex = re.sub(r'\\subsubsection\*?\{(.*?)\}', r'### \1', tex)
    
    # Formatting
    tex = re.sub(r'\\textbf\{(.*?)\}', r'**\1**', tex)
    tex = re.sub(r'\\textit\{(.*?)\}', r'*\1*', tex)
    tex = re.sub(r'\\underline\{(.*?)\}', r'<u>\1</u>', tex)
    tex = re.sub(r'\\texttt\{(.*?)\}', r'`\1`', tex)
    tex = re.sub(r'\\%', '%', tex)
    
    # Lists (Simple approach)
    tex = re.sub(r'\\begin\{itemize.*?\}', '', tex)
    tex = re.sub(r'\\end\{itemize\}', '', tex)
    tex = re.sub(r'\\begin\{enumerate.*?\}', '', tex)
    tex = re.sub(r'\\end\{enumerate\}', '', tex)
    tex = re.sub(r'\\item\s+', r'- ', tex)
    
    # Remove some environments and commands
    tex = re.sub(r'\\maketitle', '', tex)
    tex = re.sub(r'\\tableofcontents', '', tex)
    tex = re.sub(r'\\listoftables', '', tex)
    tex = re.sub(r'\\newpage', '', tex)
    tex = re.sub(r'\\addcontentsline\{.*?\}\{.*?\}\{.*?\}', '', tex)
    tex = re.sub(r'\\vspace\{.*?\}', '', tex)
    tex = re.sub(r'\\noindent\s*', '', tex)
    tex = re.sub(r'\\thispagestyle\{.*?\}', '', tex)
    
    # Tables cleanup (rough)
    tex = re.sub(r'\\begin\{table.*?\}', '', tex)
    tex = re.sub(r'\\end\{table\}', '', tex)
    tex = re.sub(r'\\centering', '', tex)
    tex = re.sub(r'\\caption\{(.*?)\}', r'**表：\1**\n', tex)
    tex = re.sub(r'\\label\{.*?\}', '', tex)
    tex = re.sub(r'\\begin\{threeparttable\}', '', tex)
    tex = re.sub(r'\\end\{threeparttable\}', '', tex)
    tex = re.sub(r'\\begin\{tablenotes\}.*?\\item (.*?)\\end\{tablenotes\}', r'\n註：\1\n', tex, flags=re.DOTALL)
    tex = re.sub(r'\\toprule', '', tex)
    tex = re.sub(r'\\midrule', '', tex)
    tex = re.sub(r'\\bottomrule', '', tex)
    tex = re.sub(r'\\small', '', tex)
    tex = re.sub(r'\\scriptsize', '', tex)
    
    # Math inline
    tex = re.sub(r'\$(.*?)\$', r'$\1$', tex) # Keep math as is
    
    # Tables (Very basic, better to let the user know tables might need manual tweaking or just keep them as raw tex or format simply)
    # Actually, let's just leave tabular environments mostly as text or remove the begin/end and keep the content.
    
    # Clean up multiple newlines
    tex = re.sub(r'\n{3,}', '\n\n', tex)
    
    return tex.strip()

def main():
    input_file = r'd:\Repositories\nsysu_finance_paper_v3\09-LaTeX論文產出與編譯\main.tex'
    output_file = r'd:\Repositories\nsysu_finance_paper_v3\09-LaTeX論文產出與編譯\main.md'
    
    with open(input_file, 'r', encoding='utf-8') as f:
        tex = f.read()
        
    md = latex_to_md(tex)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(md)
    print("Conversion complete.")

if __name__ == "__main__":
    main()
