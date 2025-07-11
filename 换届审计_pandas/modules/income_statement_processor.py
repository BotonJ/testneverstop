# /modules/income_statement_processor.py

import re
import pandas as pd
from src.utils.logger_config import logger

def _find_header_row(ws_src, headers_to_find):
    """辅助函数：查找包含指定关键字的表头行。"""
    for row in ws_src.iter_rows(min_row=1, max_row=20): # 通常表头在文件前20行
        row_values = [str(cell.value).strip() for cell in row if cell.value]
        if any(header in value for value in row_values for header in headers_to_find):
            return row[0].row
    return None

def process_income_statement(ws_src, sheet_name, mapping_df, alias_map_df):
    """
    【新版】处理单个业务活动表Sheet，提取数据并返回字典列表。
    """
    if mapping_df is None or mapping_df.empty:
        logger.warning(f"跳过Sheet '{sheet_name}' 的处理，因为'业务活动表逐行'配置为空。")
        return []

    year_match = re.search(r'(\d{4})', sheet_name)
    year = year_match.group(1) if year_match else "未知年份"
    
    # 查找表头行，以定位数据区域
    header_row_num = _find_header_row(ws_src, ['项目', '行次'])
    if not header_row_num:
        logger.error(f"在Sheet '{sheet_name}' 中未能定位到表头行，无法提取数据。")
        return []

    records = []
    
    # 将mapping配置转换为更易于查找的字典
    # key: 标准字段名, value: (本期列, 上期列)
    mapping_dict = {
        row['字段名']: (row['源期末坐标'], row['源期初坐标'])
        for _, row in mapping_df.iterrows()
    }
    
    # 构建别名 -> 标准名的反向映射
    alias_to_standard = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard = str(row['标准科目名']).strip()
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [alias.strip() for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_to_standard[alias] = standard

    # 从表头行下一行开始遍历数据
    for row in ws_src.iter_rows(min_row=header_row_num + 1, max_col=10): # 假设数据在10列以内
        subject_cell = row[0] # 通常项目名称在A列
        if not subject_cell.value:
            continue
        
        subject_name = str(subject_cell.value).strip()
        
        # 检查当前项目是否是我们关心的项目（通过别名或标准名）
        standard_name = alias_to_standard.get(subject_name, subject_name)

        if standard_name in mapping_dict:
            this_year_col, last_year_col = mapping_dict[standard_name]
            
            this_year_val = ws_src[f"{this_year_col}{subject_cell.row}"].value
            last_year_val = ws_src[f"{last_year_col}{subject_cell.row}"].value

            record = {
                "来源Sheet": sheet_name,
                "报表类型": "业务活动表",
                "年份": year,
                "项目": standard_name,
                "本期金额": this_year_val,
                "上期金额": last_year_val # “源期初坐标”对应“上期金额”
            }
            records.append(record)

    return records