# -*- coding: utf-8 -*-
import pandas as pd, glob, os
folder = "台灣碩博士論文題目清單"
for f in ["fb260621_ESG-114-01_05.csv", "113ndltd.csv"]:
    p = os.path.join(folder, f)
    for enc in ["utf-8-sig", "utf-16", "cp950", "big5"]:
        try:
            d = pd.read_csv(p, encoding=enc, nrows=3, dtype=str)
            print("====", f, "enc=", enc, "cols=", list(d.columns))
            print(d.head(2).to_string())
            break
        except Exception as e:
            print(f, enc, "FAIL", str(e)[:60])
    print()
