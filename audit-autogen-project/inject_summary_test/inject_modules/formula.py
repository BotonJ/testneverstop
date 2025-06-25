import pandas as pd
from .log_utils import log_write_text

def inject_formula_sheet(ws_tgt, mapping_file, log=None):
    try:
        df = pd.read_excel(mapping_file, sheet_name="合计公式配置")
        for _, row in df.iterrows():
            cell = str(row.get("变动单元格", "")).strip()
            formula = str(row.get("变动公式", "")).strip()
            if cell and formula:
                ws_tgt[cell] = formula
                if log is not None:                    
                    log_write_text(log, "success", cell, "", "", f"写入公式 {formula}")
    except Exception as e:
        if log is not None:
            log_write_text(log, "error", "", "", "", f"读取失败: {e}")
