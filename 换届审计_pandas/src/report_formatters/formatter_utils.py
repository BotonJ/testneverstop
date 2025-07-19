# /src/report_formatters/formatter_utils.py

import re
from typing import Dict, Optional

def create_dynamic_year_formatter(audit_period_str: Optional[str]) -> Dict[int, str]:
    """
    【核心工具函数】
    根据一个完整的审计期间字符串（例如 "2021年3月-2025年5月"），
    生成一个年份到其特定格式化表头的映射字典。
    
    这个函数是您提供的 check_year_format.py 脚本的生产版本。

    Args:
        audit_period_str (str): 从 HeaderMapping 中读取的审计期间字符串。

    Returns:
        Dict[int, str]: 一个字典，键是整数年份，值是格式化后的表头字符串。
                        例如：{2021: "2021年3-12月累计数", 2022: "2022年", ...}
    """
    year_to_header_map = {}
    if not audit_period_str or not isinstance(audit_period_str, str):
        return year_to_header_map

    # 使用正则表达式解析审计期间
    # 支持 "2021年3月-2025年5月" 和 "2021年3月至2025年5月" 两种格式
    match = re.match(r'(\d{4})年(\d{1,2})月[-至](\d{4})年(\d{1,2})月', audit_period_str.replace(" ", ""))
    if not match:
        # 如果格式不匹配，返回空字典，调用者将使用默认格式
        return year_to_header_map

    start_year, start_month, end_year, end_month = map(int, match.groups())

    # 模拟循环，为范围内的每一年生成正确的表头
    for year in range(start_year, end_year + 1):
        # 默认格式，例如 "2022年"
        year_str = f"{year}年"
        
        # 判断是否为起始年，且起始月份不是1月
        if year == start_year and start_month != 1:
            year_str = f"{year}年{start_month}-12月累计数"
            
        # 判断是否为终止年，且终止月份不是12月
        if year == end_year and end_month != 12:
            # 如果起止在同一年，格式为 "2025年1-5月累计数"
            if start_year == end_year:
                year_str = f"{year}年{start_month}-{end_month}月累计数"
            # 如果是跨年的最后一年，格式为 "2025年1-5月累计数"
            else:
                year_str = f"{year}年1-{end_month}月累计数"
                
        year_to_header_map[year] = year_str
        
    return year_to_header_map