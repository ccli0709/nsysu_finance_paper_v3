import os

filepath = "src/03_paper_generation/09_generate_latex.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(r'$D \times \Delta\text{LNREV}$', r'$DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$\text{Env\_Var} \times D \times \Delta\text{LNREV}$', r'$Env\_Var_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$\text{Env\_Var}$', r'$Env\_Var_{i,t}$')
content = content.replace(r'$SIZE \times D \times \Delta\text{LNREV}$', r'$SIZE_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$AI \times D \times \Delta\text{LNREV}$', r'$AI_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$EI \times D \times \Delta\text{LNREV}$', r'$EI_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$ROA \times D \times \Delta\text{LNREV}$', r'$ROA_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$LEV \times D \times \Delta\text{LNREV}$', r'$LEV_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')
content = content.replace(r'$SUCC\_DEC \times D \times \Delta\text{LNREV}$', r'$SUCC\_DEC_{i,t} \times DEC_{i,t} \times \Delta LNREV_{i,t}$')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
