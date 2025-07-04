import pandas as pd


def inject_formula_sheet(ws_tgt, mapping_file, log=None):
    try:
        df = pd.read_excel(mapping_file, sheet_name="合计公式配置")
        for _, row in df.iterrows():
            cell = str(row.get("变动单元格", "")).strip()
            formula = str(row.get("变动公式", "")).strip()
            if cell and formula:
                ws_tgt[cell] = formula
                
    except Exception as e:        
        return
