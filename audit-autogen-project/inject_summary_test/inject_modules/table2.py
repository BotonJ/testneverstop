from .log_utils import log_write_text
def inject_table2(wb_src, ws_tgt, conf, df_map, log=None):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        start_row = int(row["起始行"])
        end_row = int(row["终止行"])
        src_col_init = str(row["来源列（期初）"]).strip()
        src_col_final = str(row["来源列（期末）"]).strip()
        tgt_row = int(row["目标起始单元格"][1:])
        tgt_col_prefix = str(row["目标起始单元格"][0]).strip()
        skip_strs = [s.strip() for s in str(row.get("跳过行", "")).split(",") if s.strip()]
        skip_zero = str(row.get("是否跳过均为0", "")).strip() == "是"

        out_row = tgt_row
        for r in range(start_row, end_row + 1):
            subject = ws_src_init[f"A{r}"].value
            if not subject or any(skip in subject for skip in skip_strs):
                continue

            val_init = ws_src_init[f"{src_col_init}{r}"].value
            val_final = ws_src_final[f"{src_col_final}{r}"].value

            if skip_zero and ((not val_init or val_init == 0) and (not val_final or val_final == 0)):
                continue

            ws_tgt[f"{tgt_col_prefix}{out_row}"] = subject
            ws_tgt[f"{chr(ord(tgt_col_prefix)+1)}{out_row}"] = val_init or ""
            ws_tgt[f"{chr(ord(tgt_col_prefix)+2)}{out_row}"] = val_final or ""
            ws_tgt[f"{chr(ord(tgt_col_prefix)+3)}{out_row}"] = f"={chr(ord(tgt_col_prefix)+2)}{out_row}-{chr(ord(tgt_col_prefix)+1)}{out_row}"

            if log is not None:                
                log_write_text(log, "success", subject, val_init, val_final, f"写入 {tgt_col_prefix}{out_row}")

            out_row += 1
