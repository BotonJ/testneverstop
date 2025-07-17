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
    【V3.3 - 最终版】
    - 保留配置表的完整性，不再要求用户删除合计行。
    - 智能判断'聚合区块'和'单一合计项'，并正确标记'所属区块'。
    """
    logger.info(f"--- 开始处理资产负债表: '{sheet_name}' (使用V3.3'最终版'逻辑) ---")
    if blocks_df is None or blocks_df.empty: 
        logger.warning(f"'{sheet_name}': '资产负债表区块'配置为空，跳过处理。")
        return []
    
    if '合计项名称' not in blocks_df.columns:
        logger.warning(f"配置警告: '资产负债表区块' Sheet页缺少 '合计项名称' 列，内部分项核对可能不准确。")
        blocks_df['合计项名称'] = None

    alias_lookup = {}
    total_items_set = set()
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard_clean = normalize_name(row['标准科目名'])
            if not standard_clean: continue
            
            if '科目类型' in row and str(row['科目类型']).strip() == '合计':
                total_items_set.add(standard_clean)
            
            alias_lookup[standard_clean] = standard_clean
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [normalize_name(alias) for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_lookup[alias] = standard_clean

    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    for _, block_row in blocks_df.iterrows():
        block_name = block_row.get('区块名称')
        if pd.isna(block_name): continue

        start_row, search_col = _get_row_and_col_from_address(block_row['起始单元格'])
        end_row, _ = _get_row_and_col_from_address(block_row['终止单元格'])

        if not start_row or not end_row or not search_col:
            logger.warning(f"处理区块'{block_name}'时，起始/终止单元格格式不正确，已跳过。")
            continue
        
        block_tag = block_row.get('合计项名称')
        if pd.isna(block_tag):
             block_tag = block_name

        for r_idx in range(start_row, end_row + 1):
            cell_val = ws_src[f"{search_col}{r_idx}"].value
            if not cell_val: continue
            
            subject_name_clean = normalize_name(cell_val)
            if not subject_name_clean: continue

            standard_name = alias_lookup.get(subject_name_clean, subject_name_clean)
            subject_type = '合计' if standard_name in total_items_set else '普通'

            if start_row == end_row:
                subject_type = '合计'

            start_val = ws_src[f"{block_row['源期初列']}{r_idx}"].value
            end_val = ws_src[f"{block_row['源期末列']}{r_idx}"].value
            
            records.append({
                "来源Sheet": sheet_name, "报表类型": "资产负债表", "年份": year,
                "项目": standard_name,
                "所属区块": block_tag,
                "科目类型": subject_type,
                "期初金额": start_val, "期末金额": end_val
            })
            
    logger.info(f"--- 资产负债表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records