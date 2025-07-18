# /src/legacy_runner.py

import pandas as pd
import logging
from openpyxl import load_workbook as lw
from modules.balance_sheet_processor import process_balance_sheet
from modules.income_statement_processor import process_income_statement

logger = logging.getLogger(__name__)

def run_legacy_extraction(source_path: str, mapping_configs: dict):
    """
    【V3.0 - 职责单一版】
    只负责从源Excel文件中提取所有数据，并将其整合成一个干净的、
    标准化的Pandas DataFrame (raw_df)。
    """
    logger.info("  -> 开始从源文件提取数据...")
    try:
        workbook = lw(source_path, data_only=True)
    except FileNotFoundError:
        logger.error(f"源数据文件未找到: {source_path}")
        return pd.DataFrame()

    all_data = []
    
    # 获取配置
    alias_map_df = mapping_configs.get('科目等价映射')
    bs_map_df = mapping_configs.get('资产负债表区块')
    is_map_df = mapping_configs.get('业务活动表逐行')
    is_summary_config_df = mapping_configs.get('业务活动表汇总注入配置')

    for sheet_name in workbook.sheetnames:
        # 使用更稳健的关键词判断
        lower_sheet_name = sheet_name.lower()
        if '资产' in sheet_name or 'zcfz' in lower_sheet_name:
            records = process_balance_sheet(workbook[sheet_name], sheet_name, bs_map_df, alias_map_df)
            if records: all_data.extend(records)
        elif '业务' in sheet_name or 'yewu' in lower_sheet_name:
            if is_map_df is not None:
                records = process_income_statement(workbook[sheet_name], sheet_name, alias_map_df, is_map_df.to_dict('records'), is_summary_config_df)
                if records: all_data.extend(records)
    
    if not all_data:
        logger.warning("未能从源文件中提取到任何有效数据。")
        return pd.DataFrame()
        
    raw_df = pd.DataFrame(all_data)
    # 清洗年份，确保是数字
    raw_df['年份'] = pd.to_numeric(raw_df['年份'], errors='coerce')
    raw_df.dropna(subset=['年份'], inplace=True)
    raw_df['年份'] = raw_df['年份'].astype(int)

    logger.info(f"  -> 数据提取完成，共生成 {len(raw_df)} 条记录。")
    return raw_df