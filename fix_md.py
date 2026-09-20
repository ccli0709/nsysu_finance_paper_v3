import os
import glob
import re

replacements = {
    # Replace markdown equation parts
    r'dLNSGA': r'\Delta LNSGA',
    r'dLNREV': r'\Delta LNREV',
    r'Decrease': r'SUCC\_DEC',
    r'Size': r'SIZE',
    r'Lev': r'LEV',
    r'Water_Rate': r'WATER\_RATE',
    r'Water_Disc': r'WATER\_DISC',
    r'Waste_Intensity': r'WASTE\_INTENSITY',
    r'Waste_Density': r'WASTE\_DENSITY',
    r'Waste_Fine': r'WASTE\_FINE',
    r'Waste_FineInt': r'WASTE\_FINE\_INT',
}

md_files = glob.glob("docs/**/*.md", recursive=True)
for filepath in md_files:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    for old, new in replacements.items():
        content = re.sub(rf'\b{old}\b', new, content)
    
    # Specific equation replacements
    # Find $\Delta LNSGA = \beta_0 + \beta_1 ...$ and add i,t
    content = content.replace(r'\Delta LNSGA =', r'\Delta LNSGA_{i,t} =')
    content = content.replace(r'\Delta LNREV', r'\Delta LNREV_{i,t}')
    content = content.replace(r' D \times', r' DEC_{i,t} \times')
    content = content.replace(r'(D \times', r'(DEC_{i,t} \times')
    content = content.replace(r' D ', r' DEC_{i,t} ')
    content = content.replace(r'DEC \times', r'DEC_{i,t} \times')
    
    # adding i,t to all uppercase words in equations? No, too risky.
    # We will just append _{i,t} to SIZE, LEV, ROA, AI, EI inside $...$ blocks.
    def add_subscripts(m):
        eq = m.group(0)
        # Add a proper raw string regex or replace
        for var in ['SIZE', 'LEV', 'ROA', 'AI', 'EI', r'SUCC\_DEC', r'WATER\_RATE', r'WATER\_DISC', r'WASTE\_INTENSITY', r'WASTE\_DENSITY', r'WASTE\_FINE', r'WASTE\_FINE\_INT']:
            # avoid double subscript if already done
            eq = re.sub(rf'{var}(?!_)', f"{var}_{{i,t}}", eq)
        # Ensure error term
        if '=' in eq and r'\varepsilon' not in eq:
            eq = eq.replace(r'$', r' + \varepsilon_{i,t}$')
        return eq

    content = re.sub(r'\$.*?\$', add_subscripts, content)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Updated {filepath}")
