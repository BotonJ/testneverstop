import pandas as pd
from .log_utils import log_write_text
def inject_table3(wb_src, ws_tgt, conf, df_map, log=None):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        if pd.isna(row.get("来源字段")):
            continue

        src_init_cell = str(row.get("来源单元格（期初）", "")).strip()
        src_final_cell = str(row.get("来源单元格（期末）", "")).strip()
        tgt_init_cell = str(row.get("目标单元格（期初）", "")).strip()
        tgt_final_cell = str(row.get("目标单元格（期末）", "")).strip()
        add_cell = str(row.get("增加单元格", "")).strip()
        sub_cell = str(row.get("减少单元格", "")).strip()

        val_init = ws_src_init[src_init_cell].value if src_init_cell else 0
        val_final = ws_src_final[src_final_cell].value if src_final_cell else 0

        if tgt_init_cell:
            ws_tgt[tgt_init_cell] = val_init
        if tgt_final_cell:
            ws_tgt[tgt_final_cell] = val_final

        try:
            num_init = float(val_init) if val_init not in [None, ""] else 0
            num_final = float(val_final) if val_final not in [None, ""] else 0
            diff = num_final - num_init
            if diff > 0 and add_cell:
                ws_tgt[add_cell] = diff
            elif diff <= 0 and sub_cell:
                ws_tgt[sub_cell] = abs(diff)

            if log is not None:                
                log_write_text(log, "success",src_init_cell, val_init, val_final, f"写入 {tgt_init_cell}/{tgt_final_cell}")

        except Exception as e:
            if log is not None:                
                log_write_text(log, "error", src_init_cell, val_init, val_final, f"净资产差值写入失败: {e}")
