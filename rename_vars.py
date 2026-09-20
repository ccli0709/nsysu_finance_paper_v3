import os
import glob
import re

replacements = {
    "Water_Rate": "WATER_RATE",
    "Water_Disc": "WATER_DISC",
    "Water_Intensity": "WATER_INTENSITY",
    "Waste_Intensity": "WASTE_INTENSITY",
    "Waste_Density": "WASTE_DENSITY",
    "Waste_Fine": "WASTE_FINE",
    "Waste_FineInt": "WASTE_FINE_INT",
    "Decrease": "SUCC_DEC",
    "Lev": "LEV",
    "Size": "SIZE",
    "dLNSGA": "dLNSGA",
    "dLNREV": "dLNREV",
}

py_files = glob.glob("src/**/*.py", recursive=True)
for filepath in py_files:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # We must be careful not to replace substrings indiscriminately, but these are pretty unique
    # We'll replace exact string matches or standalone words.
    
    for old, new in replacements.items():
        # Replace string literal exact matches and dictionary keys
        content = re.sub(rf'\b{old}\b', new, content)
        
    # Specifically change D to DEC
    # D is too short, we should only replace it if it's a quote or a column name
    # like "D" or df["D"] or 'D'
    content = re.sub(r'\"D\"', '"DEC"', content)
    content = re.sub(r'\'D\'', "'DEC'", content)
    # in variables like D_x_dREV
    content = re.sub(r'\bD_x_dREV\b', 'DEC_x_dREV', content)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Updated {filepath}")
