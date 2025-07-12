# /src/legacy_runner.py
import re
import sys
import os
import pandas as pd
from openpyxl import load_workbook

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(PROJECT_ROOT)

from src.utils.logger_config import logger
from modules.mapping_loader import load_mapping_file
from modules.balance_sheet_processor import process_balance_sheet
from modules.income_statement_processor import process_income_statement

def run_legacy_extraction(source_path, mapping_path):
    """
    【最终修复版 - 总指挥官】
    整合所有正确逻辑，健壮地处理各类Sheet名。
    """
    logger.info("--- 开始执行【最终修复版】数据提取流程 ---")
    
    mapping = load_mapping_file(mapping_path)
    if not mapping:
        logger.error("因映射文件加载失败，数据提取流程终止。")
        return None
        
    blocks_df = mapping.get("blocks_df")
    alias_map_df = mapping.get("alias_map_df")
    yewu_line_map = mapping.get("yewu_line_map")

    try:
        wb_src = load_workbook(source_path, data_only=True)
    except FileNotFoundError:
        logger.error(f"源数据文件未找到: {source_path}")
        return None

    all_records = []
    
    # --- 主循环：采用最终的、最健壮的识别逻辑 ---
    for original_sheet_name in wb_src.sheetnames:
        ws_src = wb_src[original_sheet_name]
        sheet_name = original_sheet_name.strip()

        if ws_src.sheet_state == 'hidden':
            logger.warning(f"跳过隐藏的Sheet: '{sheet_name}'")
            continue

        # 使用正则表达式智能判断Sheet类型
        is_balance_sheet = re.search(r'(\d{4})\s*年?\s*资产负债表', sheet_name, re.IGNORECASE) or \
                           re.search(r'^(\d{4})\s*z$', sheet_name, re.IGNORECASE)
                           
        is_income_statement = re.search(r'(\d{4})\s*年?\s*业务活动表', sheet_name, re.IGNORECASE) or \
                              re.search(r'^(\d{4})\s*y$', sheet_name, re.IGNORECASE)

        if is_balance_sheet:
            balance_sheet_records = process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df)
            if balance_sheet_records:
                all_records.extend(balance_sheet_records)

        elif is_income_statement:
            income_statement_records = process_income_statement(ws_src, sheet_name, yewu_line_map, alias_map_df)
            if income_statement_records:
                all_records.extend(income_statement_records)
        else:
             logger.warning(f"跳过Sheet: '{sheet_name}'，因其命名不符合任何已知规则。")


    if not all_records:
        logger.error("未能从源文件中提取到任何有效数据记录。请检查soce.xlsx中的Sheet名是否符合规则（如 '2019资产负债表' 或 '2019Z'）。")
        return pd.DataFrame()

    final_df = pd.DataFrame(all_records)

    amount_cols = ['期初金额', '期末金额', '本期金额', '上期金额']
    for col in amount_cols:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)

    logger.info(f"--- 数据提取流程结束，成功生成包含 {len(final_df)} 条记录的DataFrame。---")
    return final_df