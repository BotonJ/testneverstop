# /modules/balance_sheet_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger
from modules.utils import normalize_name

def _get_row_and_col_from_address(address):
    """从单元格地址（如'A13'）中提取行号和列字母。"""
    if not address or not isinstance(address, str):
        return None, None
    match = re.match(r"([A-Z]+)(\d+)", str(address).strip())
    if match:
        col, row = match.groups()
        return int(row), col
    return None, None

def process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df):
    """
    【最终版 V3 - 智能推断】
    1. 不再需要'科目搜索列'，改为从'起始单元格'动态推断。
    2. 严格按照“区块”处理，为每条数据打上正确的“所属区块”标签。
    """
    logger.info(f"--- 开始处理资产负债表: '{sheet_name}' (使用最终版'智能推断'逻辑) ---")
    if blocks_df is None or blocks_df.empty: 
        logger.warning(f"'{sheet_name}': '资产负债表区块'配置为空，跳过处理。")
        return []

    alias_lookup = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard_clean = normalize_name(row['标准科目名'])
            if not standard_clean: continue
            
            subj_type = '合计' if '科目类型' in row and str(row['科目类型']).strip() == '合计' else '普通'
            alias_lookup[standard_clean] = (standard_clean, subj_type)
            
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [normalize_name(alias) for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_lookup[alias] = (standard_clean, subj_type)

    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    for _, block_row in blocks_df.iterrows():
        block_name = block_row.get('区块名称')
        if pd.isna(block_name): continue

        start_row, search_col = _get_row_and_col_from_address(block_row['起始单元格'])
        end_row, _ = _get_row_and_col_from_address(block_row['终止单元格'])

        if not start_row or not end_row or not search_col:
            logger.warning(f"处理区块'{block_name}'时，起始/终止单元格格式不正确或无法提取搜索列，已跳过。")
            continue

        logger.debug(f"处理区块'{block_name}': 在'{search_col}'列, 扫描行 {start_row}-{end_row}")

        for r_idx in range(start_row, end_row + 1):
            cell_val = ws_src[f"{search_col}{r_idx}"].value
            if not cell_val: continue
            
            subject_name_clean = normalize_name(cell_val)
            if not subject_name_clean: continue

            if subject_name_clean in alias_lookup:
                standard_name, subject_type = alias_lookup[subject_name_clean]
            else:
                standard_name, subject_type = subject_name_clean, '普通'

            start_val_col, end_val_col = block_row['源期初列'], block_row['源期末列']
            start_val = ws_src[f"{start_val_col}{r_idx}"].value
            end_val = ws_src[f"{end_val_col}{r_idx}"].value

            records.append({
                "来源Sheet": sheet_name, "报表类型": "资产负债表", "年份": year,
                "项目": standard_name,
                "所属区块": block_name, 
                "科目类型": subject_type,
                "期初金额": start_val, "期末金额": end_val
            })
            
    logger.info(f"--- 资产负债表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records