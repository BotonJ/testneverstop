# /modules/balance_sheet_processor.py

import re
import pandas as pd
# 统一使用新的日志记录器
from src.utils.logger_config import logger

def _find_subject_row(ws_src, standard_name, aliases, search_cols=['A', 'E']):
    """辅助函数：在源Sheet的指定列中查找科目所在的行。"""
    all_names_to_find = [standard_name] + aliases
    for row in ws_src.iter_rows(min_row=1, max_row=ws_src.max_row):
        for col_letter in search_cols:
            cell = ws_src[f"{col_letter}{row[0].row}"]
            if cell.value and str(cell.value).strip() in all_names_to_find:
                return cell.row
    return None

def process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df):
    """
    【新版】处理单个资产负债表Sheet，提取数据并返回字典列表。
    """
    if blocks_df is None or blocks_df.empty:
        logger.warning(f"跳过Sheet '{sheet_name}' 的处理，因为'资产负债表区块'配置为空。")
        return []

    year_match = re.search(r'(\d{4})', sheet_name)
    year = year_match.group(1) if year_match else "未知年份"

    records = []
    
    alias_to_standard = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard = str(row['标准科目名']).strip()
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [alias.strip() for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_to_standard[alias] = standard

    for _, block_row in blocks_df.iterrows():
        standard_name = block_row['区块名称']
        
        # --- vvvvvvvv 核心BUG修复 vvvvvvvv ---
        # 如果'区块名称'这一列是空的 (Pandas读取空单元格为NaN), 就跳过这一整行
        if pd.isna(standard_name):
            continue
        # --- ^^^^^^^^ 核心BUG修复 ^^^^^^^^ ---
        
        start_col = block_row['源期初列'] # 现在这一行安全了
        end_col = block_row['源期末列']

        aliases = [alias for alias, std in alias_to_standard.items() if std == standard_name]
        
        found_row = _find_subject_row(ws_src, standard_name, aliases, search_cols=['A', 'E'])

        if found_row:
            start_val = ws_src[f"{start_col}{found_row}"].value
            end_val = ws_src[f"{end_col}{found_row}"].value
            record = {
                "来源Sheet": sheet_name, "报表类型": "资产负债表", "年份": year,
                "项目": standard_name, "期初金额": start_val, "期末金额": end_val
            }
            records.append(record)
        else:
            logger.warning(f"在Sheet '{sheet_name}' 中未找到项目 '{standard_name}' 或其任何别名。")

    return records