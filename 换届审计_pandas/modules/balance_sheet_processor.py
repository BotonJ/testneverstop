# /modules/balance_sheet_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger
from modules.utils import normalize_name # <-- 导入我们新的净化函数

def process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df):
    """
    【最终版 V2 - 忠于经典逻辑】
    增加了 normalize_name 清洗，确保匹配的健壮性。
    """
    logger.info(f"--- 开始处理资产负债表: '{sheet_name}' (使用带清洗的'全局扫描'逻辑) ---")

    alias_lookup = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard = normalize_name(row['标准科目名']) # 清洗标准名
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [normalize_name(alias) for alias in str(row[col]).split(',')] # 清洗别名
                    for alias in aliases:
                        if alias: alias_lookup[alias] = standard

    src_dict = {}
    for i in range(1, ws_src.max_row + 1):
        name_a_raw = ws_src[f"A{i}"].value
        if name_a_raw:
            name_a_clean = normalize_name(name_a_raw)
            if name_a_clean:
                name_std = alias_lookup.get(name_a_clean, name_a_clean)
                src_dict[name_std] = {"期初": ws_src[f"C{i}"].value, "期末": ws_src[f"D{i}"].value}

        name_e_raw = ws_src[f"E{i}"].value
        if name_e_raw:
            name_e_clean = normalize_name(name_e_raw)
            if name_e_clean:
                name_std = alias_lookup.get(name_e_clean, name_e_clean)
                if name_std not in src_dict:
                     src_dict[name_std] = {"期初": ws_src[f"G{i}"].value, "期末": ws_src[f"H{i}"].value}
    
    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    for subject_name, values in src_dict.items():
        # 这里我们也使用清洗后的标准名进行比较
        total_subjects_clean = {normalize_name(s) for s in ['资产总计', '负债合计', '净资产合计', '流动资产合计', '非流动资产合计', '流动负债合计', '非流动负债合计']}
        subject_type = '合计' if subject_name in total_subjects_clean else '普通'

        records.append({
            "来源Sheet": sheet_name, "报表类型": "资产负债表", "年份": year,
            "项目": subject_name, "科目类型": subject_type,
            "期初金额": values["期初"], "期末金额": values["期末"]
        })
        
    logger.info(f"--- 资产负债表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records