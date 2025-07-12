# /modules/balance_sheet_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger

def process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df):
    """
    【回溯版 - 忠于原始逻辑】
    模拟 fill_balance_anchor.py 的“全局扫描，字典匹配”算法。
    """
    logger.info(f"--- 开始处理资产负债表: '{sheet_name}' (使用'全局扫描'逻辑) ---")

    alias_lookup = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard = str(row['标准科目名']).strip()
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [alias.strip() for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_lookup[alias] = standard

    src_dict = {}
    for i in range(1, ws_src.max_row + 1):
        name_a = ws_src[f"A{i}"].value
        if name_a and str(name_a).strip():
            name_std = alias_lookup.get(str(name_a).strip(), str(name_a).strip())
            src_dict[name_std] = {"期初": ws_src[f"C{i}"].value, "期末": ws_src[f"D{i}"].value}

        name_e = ws_src[f"E{i}"].value
        if name_e and str(name_e).strip():
            name_std = alias_lookup.get(str(name_e).strip(), str(name_e).strip())
            if name_std not in src_dict:
                 src_dict[name_std] = {"期初": ws_src[f"G{i}"].value, "期末": ws_src[f"H{i}"].value}
    
    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    for subject_name, values in src_dict.items():
        total_subjects = ['资产总计', '负债合计', '净资产合计', '流动资产合计', '非流动资产合计', '流动负债合计', '非流动负债合计']
        subject_type = '合计' if subject_name in total_subjects else '普通'

        records.append({
            "来源Sheet": sheet_name, "报表类型": "资产负债表", "年份": year,
            "项目": subject_name, "科目类型": subject_type,
            "期初金额": values["期初"], "期末金额": values["期末"]
        })
        
    logger.info(f"--- 资产负债表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records