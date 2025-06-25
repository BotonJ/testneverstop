from .log_utils import log_write_text
def inject_table1(wb_src, ws_tgt, conf, df_map, log=None):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        src_field = str(row["来源字段"]).strip()
        tgt_init_cell = str(row["目标单元格（期初）"]).strip()
        tgt_final_cell = str(row["目标单元格（期末）"]).strip()
        var_cell = str(row.get("变动单元格", "")).strip()
        var_formula = str(row.get("变动公式", "")).strip()

        val_init, val_final = None, None
        for i in range(1, 100):
            name = ws_src_init[f"A{i}"].value
            if name and src_field in str(name):
                val_init = ws_src_init[f"B{i}"].value
                break
        for i in range(1, 100):
            name = ws_src_final[f"A{i}"].value
            if name and src_field in str(name):
                val_final = ws_src_final[f"C{i}"].value
                break

        ws_tgt[tgt_init_cell] = val_init
        ws_tgt[tgt_final_cell] = val_final

        if var_cell and var_formula:
            ws_tgt[var_cell] = var_formula

        if log is not None:            
            log_write_text(log, "success",  src_field, val_init, val_final, f"写入 {tgt_init_cell}/{tgt_final_cell}")
