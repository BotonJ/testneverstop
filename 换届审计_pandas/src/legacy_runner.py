# /src/legacy_runner.py

import pandas as pd
import openpyxl
import logging
from typing import Dict, List

# 导入您项目中的实际处理模块
from modules.balance_sheet_processor import process_balance_sheet
from modules.income_statement_processor import process_income_statement

logger = logging.getLogger(__name__)

def run_legacy_extraction(source_path: str, mapping_configs: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    【最终修复版】
    根据每个子处理函数的精确函数签名（参数列表），传递正确的参数。
    """
    logger.info("--- 开始执行【最终修复版 V4】数据提取流程 ---")
    
    if not mapping_configs:
        logger.error("传入的mapping_configs为空，无法继续提取。")
        return pd.DataFrame()

    try:
        # 从总配置字典中，提前取出所有子模块可能需要的配置表
        alias_map_df = mapping_configs['科目等价映射']
        # header_map_df is no longer needed by process_income_statement
        # header_map_df = mapping_configs['HeaderMapping']
        bs_map_df = mapping_configs['资产负债表区块']
        is_map_df = mapping_configs['业务活动表逐行']
    except KeyError as e:
        available_keys = list(mapping_configs.keys())
        logger.error(f"配置文件'mapping_file.xlsx'中缺少关键Sheet: {e}，无法继续。")
        logger.error(f"程序实际从Excel文件中读取到的Sheet名列表为: {available_keys}")
        logger.error("请检查Excel中的Sheet名与代码中所需名称是否完全一致（包括空格）。")
        return pd.DataFrame()

    all_data = []
    
    try:
        workbook = openpyxl.load_workbook(source_path, data_only=True)
        all_sheet_names = workbook.sheetnames

        # --- Pass 1/2: 处理所有资产负债表 ---
        logger.info("--- [Pass 1/2] 正在处理所有资产负债表... ---")
        for sheet_name in all_sheet_names:
            if '资产' in sheet_name or 'zcfz' in sheet_name.lower():
                sheet_records = process_balance_sheet(
                    ws_src=workbook[sheet_name],
                    sheet_name=sheet_name,
                    blocks_df=bs_map_df,
                    alias_map_df=alias_map_df
                )
                if sheet_records:
                    all_data.append(pd.DataFrame(sheet_records))

        # --- Pass 2/2: 处理所有业务活动表 ---
        logger.info("--- [Pass 2/2] 正在处理所有业务活动表... ---")
        for sheet_name in all_sheet_names:
            if '业务' in sheet_name or 'yewu' in sheet_name.lower() or sheet_name.lower().endswith('y'):
                # --- [BUG修复] 确保调用的参数名与函数定义完全匹配 ---
                # 1. 'sheet' -> 'ws_src'
                # 2. 'is_map_df' -> 'yewu_line_map'
                # 3. 移除了函数不再需要的 'header_map_df'
                sheet_records = process_income_statement(
                    ws_src=workbook[sheet_name],
                    sheet_name=sheet_name,
                    yewu_line_map=is_map_df.to_dict('records'), # The function expects a list of dicts
                    alias_map_df=alias_map_df
                )
                if sheet_records:
                    all_data.append(pd.DataFrame(sheet_records))

        if not all_data:
            logger.warning("未能从任何Sheet中提取到有效数据。")
            return pd.DataFrame()

        final_df = pd.concat(all_data, ignore_index=True)
        logger.info(f"--- 数据提取流程结束，成功生成包含 {len(final_df)} 条记录的DataFrame。 ---")
        return final_df

    except FileNotFoundError:
        logger.error(f"源文件未找到: {source_path}")
        return pd.DataFrame()
    except Exception as e:
        logger.error(f"处理源文件时发生未知错误: {e}", exc_info=True)
        return pd.DataFrame()