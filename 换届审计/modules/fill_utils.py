import logging
from modules.match_utils import match_subject_name
def fill_balance_block(ws_src, ws_tgt, blocks, alias_map):
    """
    将资产负债表一个区块的数据从源表写入模板
    :param ws_src: openpyxl 的源数据 worksheet
    :param ws_tgt: openpyxl 的模板 worksheet
    :param block_conf: mapping_loader["blocks"][区块名]
    :param alias_map: mapping_loader["subject_alias_map"]
    """
    for block_name, block_conf in blocks.items():        
        for row in ws_tgt.iter_rows(min_row=block_conf["start_row"], max_row=block_conf["end_row"]):
            tgt_row = row[0].row
            if tgt_row in block_conf.get("skip_rows", []):
                continue
            src_init_col = block_conf.get("sorce_col_initial")
            src_final_col = block_conf.get("sorce_col_final")
            tgt_init_col = block_conf.get("target_col_initial")
            tgt_final_col = block_conf.get("target_col_final")
            if None in [src_init_col, src_final_col, tgt_init_col, tgt_final_col]:
                #logging.warning(f"⚠️ 缺少列配置信息，跳过区块: {block_name}")
                continue
            subject_cell = row[0].value
            if not subject_cell or not isinstance(subject_cell, str):
                continue

            candidate_names = match_subject_name(subject_cell.strip(), alias_map)

            # 遍历源数据查找匹配行
            matched = False
            for src_row in ws_src.iter_rows(min_row=2):
                src_subject = src_row[0].value
                if src_subject and src_subject.strip() in candidate_names:
                    row[block_conf["target_col_initial"] - 1].value = src_row[block_conf["sorce_col_initial"] - 1].value
                    row[block_conf["target_col_final"] - 1].value = src_row[block_conf["sorce_col_final"] - 1].value
                    matched = True
                    break  # 找到即写入并跳出
            if not matched:
                # 可选：日志记录未匹配项
                pass
