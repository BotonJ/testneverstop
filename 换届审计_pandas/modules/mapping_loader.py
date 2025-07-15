# /modules/mapping_loader.py

import pandas as pd
import openpyxl
import logging
from typing import Dict

logger = logging.getLogger(__name__)

def load_mapping_file(path: str) -> Dict[str, pd.DataFrame]:
    """
    【全新修复版】
    读取指定的mapping_file.xlsx文件，并将其所有Sheet页加载到一个字典中。
    字典的键（key）将是Excel中真实的Sheet页名称。

    Args:
        path (str): mapping_file.xlsx 的文件路径。

    Returns:
        Dict[str, pd.DataFrame]: 一个以Sheet名为键，以DataFrame为值的配置字典。
                                 如果文件不存在或为空，返回一个空字典。
    """
    try:
        # 使用 openpyxl 先获取所有真实的sheet名称
        workbook = openpyxl.load_workbook(path, read_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()

        if not sheet_names:
            logger.error(f"配置文件 '{path}' 中未发现任何Sheet页。")
            return {}

        # 使用 pandas.read_excel 一次性读取所有sheet
        # 设置 sheet_name=None 会返回一个以sheet名为键的字典
        mapping_configs = pd.read_excel(path, sheet_name=None)
        
        logger.info(f"成功从 '{path}' 加载了 {len(mapping_configs)} 个配置Sheet: {list(mapping_configs.keys())}")
        
        return mapping_configs

    except FileNotFoundError:
        logger.error(f"指定的配置文件未找到: {path}")
        return {}
    except Exception as e:
        logger.error(f"读取配置文件 '{path}' 时发生未知错误: {e}", exc_info=True)
        return {}

# --- 保留旧函数以防万一，但新流程不再使用它们 ---
def get_col_index(df, col_name):
    try:
        return df.columns.get_loc(col_name) + 1
    except KeyError:
        return None

def parse_skip_rows(skip_str):
    if not skip_str or pd.isna(skip_str):
        return []
    return [int(x.strip()) for x in str(skip_str).split(',') if x.strip()]

def load_full_mapping_as_df(path):
    # 这个函数可能在其他地方仍有使用，暂时保留
    try:
        return pd.read_excel(path, sheet_name=None)
    except FileNotFoundError:
        return None
