import os
import glob

replacements = {
    "Water_Rate_w": "WATER_RATE_w",
    "Water_Disc_GRI": "WASTE_DISC_GRI", # wait, Water or Waste? It was Water_Disc_GRI
    "Waste_Intensity_w": "WASTE_INTENSITY_w",
    "Waste_Fine_w": "WASTE_FINE_w",
    "Waste_FineInt_w": "WASTE_FINE_INT_w",
    "Water_Intensity_w": "WATER_INTENSITY_w",
    "Size_w": "SIZE_w",
    "Lev_w": "LEV_w",
    "AI_w": "AI_w",
    "EI_w": "EI_w",
    "ROA_w": "ROA_w",
    "Decrease_w": "SUCC_DEC_w",
    
    # other missing ones
    "WaterRate": "WATER_RATE",
    "WaterDisc": "WATER_DISC",
    "WaterDiscGRI": "WATER_DISC_GRI",
    
    # also for columns in df
    "Size": "SIZE",
    "Lev": "LEV",
    "Decrease": "SUCC_DEC",
    "Water_Rate": "WATER_RATE",
    "Water_Disc": "WATER_DISC",
    "Waste_Disc": "WASTE_DISC",
}

py_files = glob.glob("src/**/*.py", recursive=True)
for filepath in py_files:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    for old, new in replacements.items():
        content = content.replace(old, new)
        
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Updated {filepath}")
