# /modules/utils.py

import re
import pandas as pd
from typing import List

def normalize_name(name: str) -> str:
    """
    清理和标准化科目名称：
    - 移除所有空格（包括全角和半角）
    - 移除括号及其内容（全角和半角）
    - 转换为小写
    """
    if not isinstance(name, str):
        return ""
    
    # 移除括号及其内容
    name = re.sub(r'[\(（].*?[\)）]', '', name)
    # 移除所有类型的空格
    name = re.sub(r'\s+', '', name)
    
    return name.strip().lower()

def _get_value_from_coord(sheet, coord: str):
    """
    【从您的旧版脚本中复用】
    根据单元格坐标（如'F12'）安全地从工作表中获取值。
    如果坐标无效或单元格为空，则返回None。
    
    Args:
        sheet: openpyxl的worksheet对象。
        coord (str): 单元格坐标字符串。

    Returns:
        The cell value or None.
    """
    if not coord or not isinstance(coord, str) or pd.isna(coord):
        return None
    try:
        return sheet[coord.strip()].value
    except (KeyError, ValueError):
        # 如果坐标无效（例如，格式不正确），也返回None
        return None

# --- 为了兼容性，保留旧的函数定义 ---

def get_col_index(df, col_name):
    try:
        return df.columns.get_loc(col_name) + 1
    except KeyError:
        return None

def parse_skip_rows(skip_str) -> List[int]:
    if not skip_str or pd.isna(skip_str):
        return []
    return [int(x.strip()) for x in str(skip_str).split(',') if x.strip()]

