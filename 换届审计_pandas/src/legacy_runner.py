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
    【最终修复版 V4 - 总指挥官】
    修复了AttributeError，采用分步判断逻辑，确保健壮性。
    """
    logger.info("--- 开始执行【最终修复版 V4】数据提取流程 ---")
    
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
    processed_balance_sheets = {} 

    # --- 第一遍循环：只处理资产负债表 ---
    logger.info("--- [Pass 1/2] 正在处理所有资产负债表... ---")
    for original_sheet_name in wb_src.sheetnames:
        ws_src = wb_src[original_sheet_name]
        sheet_name = original_sheet_name.strip()

        if ws_src.sheet_state == 'hidden': continue

        # --- 核心修复：分步判断逻辑 ---
        match = re.search(r'(\d{4})', sheet_name)
        if match:
            year = match.group(1)
            # 判断是否为资产负债表
            if "资产负债表" in sheet_name or sheet_name.lower().endswith('z'):
                balance_sheet_records = process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df)
                if balance_sheet_records:
                    all_records.extend(balance_sheet_records)
                    df_temp = pd.DataFrame(balance_sheet_records)
                    
                    # 增加健壮性检查，确保'项目'列存在
                    if '项目' in df_temp.columns:
                        processed_balance_sheets[year] = {
                            "期初净资产": pd.to_numeric(df_temp.loc[df_temp['项目'] == '净资产合计', '期初金额'].sum(), errors='coerce'),
                            "期末净资产": pd.to_numeric(df_temp.loc[df_temp['项目'] == '净资产合计', '期末金额'].sum(), errors='coerce')
                        }

    # --- 第二遍循环：只处理业务活动表 ---
    logger.info("--- [Pass 2/2] 正在处理所有业务活动表... ---")
    for original_sheet_name in wb_src.sheetnames:
        ws_src = wb_src[original_sheet_name]
        sheet_name = original_sheet_name.strip()

        if ws_src.sheet_state == 'hidden': continue
        
        match = re.search(r'(\d{4})', sheet_name)
        if match:
            year = match.group(1)
            # 判断是否为业务活动表
            if "业务活动表" in sheet_name or sheet_name.lower().endswith('y'):
                net_asset_fallback = processed_balance_sheets.get(year)
                income_statement_records = process_income_statement(
                    ws_src, sheet_name, yewu_line_map, alias_map_df, net_asset_fallback
                )
                if income_statement_records:
                    all_records.extend(income_statement_records)

    if not all_records:
        logger.error("未能从源文件中提取到任何有效数据记录。")
        return pd.DataFrame()

    final_df = pd.DataFrame(all_records)

    amount_cols = ['期初金额', '期末金额', '本期金额', '上期金额']
    for col in amount_cols:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)

    logger.info(f"--- 数据提取流程结束，成功生成包含 {len(final_df)} 条记录的DataFrame。---")
    return final_df