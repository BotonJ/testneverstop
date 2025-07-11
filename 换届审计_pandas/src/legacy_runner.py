import os
import sys
# 获取当前脚本文件(legacy_runner.py)的绝对路径
# os.path.abspath(__file__) -> /path/to/your/project/src/legacy_runner.py
# os.path.dirname(...) -> /path/to/your/project/src
# os.path.dirname(...) -> /path/to/your/project/  <-- 这就是我们的项目根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from openpyxl import load_workbook
from modules.mapping_loader import load_mapping_file, load_full_mapping_as_df
from modules.balance_sheet_processor import process_balance_sheet
from modules.income_statement_processor import process_income_statement
from src.utils.logger_config import logger

def run_legacy_extraction(source_path, mapping_path):
    """
    【新版核心函数】
    此函数负责从源Excel文件中提取所有相关数据，
    并将其统一处理成一个Pandas DataFrame返回。
    它不再创建或写入任何中间Excel文件。

    Args:
        source_path (str): 源数据文件路径 (soce.xlsx)。
        mapping_path (str): 映射文件路径 (mapping_file.xlsx)。

    Returns:
        pandas.DataFrame: 一个包含所有从源文件提取和处理后的结构化数据表。
                          返回None表示提取失败。
    """
    logger.info("--- 开始执行【新版】数据提取流程 (legacy_runner) ---")
    
    try:
        # 1. 加载源数据和映射文件
        wb_src = load_workbook(source_path, data_only=True)
        # 我们现在一次性加载所有需要的映射表为DataFrame，方便后续使用
        all_mappings = load_full_mapping_as_df(mapping_path)
        alias_map_df = all_mappings.get('科目等价映射')

        logger.info("成功加载源文件和所有映射配置。")

        # 2. 初始化一个列表，用于收集所有提取到的数据
        all_data_records = []

        # 3. 遍历源文件中的所有Sheet进行处理
        for sheet_name in wb_src.sheetnames:
            ws_src = wb_src[sheet_name]
            logger.info(f"正在处理Sheet: '{sheet_name}'...")

            # ---- 资产负债表处理 ----
            if "资产负债表" in sheet_name:
                # 调用重构后的模块，它将直接返回一个数据列表
                balance_sheet_data = process_balance_sheet(
                    ws_src, 
                    sheet_name, 
                    all_mappings.get('资产负债表区块'), 
                    alias_map_df
                )
                if balance_sheet_data:
                    all_data_records.extend(balance_sheet_data)
                    logger.info(f"从 '{sheet_name}' 提取了 {len(balance_sheet_data)} 条资产负债表记录。")

            # ---- 业务活动表（收入、费用）处理 ----
            elif "业务活动表" in sheet_name:
                 # 调用重构后的模块，它将直接返回一个数据列表
                income_statement_data = process_income_statement(
                    ws_src,
                    sheet_name,
                    all_mappings.get('业务活动表逐行'),
                    alias_map_df
                )
                if income_statement_data:
                    all_data_records.extend(income_statement_data)
                    logger.info(f"从 '{sheet_name}' 提取了 {len(income_statement_data)} 条业务活动表记录。")
            else:
                logger.warning(f"跳过Sheet: '{sheet_name}'，因为它不包含指定的关键字。")

        # 4. 【核心步骤】将收集到的所有数据记录转换为Pandas DataFrame
        if not all_data_records:
            logger.error("未能从源文件中提取到任何数据记录。")
            return None
        
        logger.info(f"数据提取完成，共收集到 {len(all_data_records)} 条记录。正在转换为DataFrame...")
        
        # 将字典列表直接转换为DataFrame
        final_df = pd.DataFrame(all_data_records)

        # 5. （可选但推荐）对DataFrame进行初步的数据清洗和类型转换
        # 例如，将金额列统一转换为数值类型，无法转换的填充为0
        amount_cols = ['期初金额', '期末金额', '本期金额', '上期金额']
        for col in amount_cols:
            if col in final_df.columns:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)
        
        logger.info("--- 数据提取流程结束，成功生成DataFrame。---")

        # 6. 返回最终的DataFrame，不再保存任何文件
        return final_df

    except FileNotFoundError:
        logger.error(f"错误：找不到源文件或映射文件。请检查路径 {source_path} 和 {mapping_path}")
        return None
    except Exception as e:
        logger.error(f"在数据提取过程中发生未知错误: {e}")
        return None

# 你可以在这里添加一个测试块，以便独立运行和调试这个脚本
if __name__ == '__main__':
    # 使用相对路径，假设你从项目根目录运行
    src_file = os.path.join(PROJECT_ROOT, 'data', 'soce.xlsx')
    map_file = os.path.join(PROJECT_ROOT, 'data', 'mapping_file.xlsx')
    
    print("正在以独立模式测试 legacy_runner.py...")
    extracted_df = run_legacy_extraction(src_file, map_file)
    
    if extracted_df is not None:
        print("\n✅ 数据提取成功！")
        print("生成的DataFrame信息：")
        extracted_df.info()
        print("\nDataFrame内容预览 (前5行):")
        print(extracted_df.head())
        print("\nDataFrame内容预览 (后5行):")
        print(extracted_df.tail())
    else:
        print("\n❌ 数据提取失败。")