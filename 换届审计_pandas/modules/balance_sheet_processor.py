# /modules/balance_sheet_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger

def process_balance_sheet(ws_src, sheet_name, blocks_df, alias_map_df):
    """
    【最终版 - 忠于原始逻辑】
    模拟 fill_balance_anchor.py 的“全局扫描，字典匹配”算法。
    """
    logger.info(f"--- 开始处理资产负债表: '{sheet_name}' (使用原始'全局扫描'逻辑) ---")

    # --- 1. 构建别名->标准名查找字典 ---
    alias_lookup = {}
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard = str(row['标准科目名']).strip()
            # 将所有别名都指向标准名
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [alias.strip() for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias:
                            alias_lookup[alias] = standard

    # --- 2. 全局扫描源Sheet，构建数据“电话本” (src_dict) ---
    src_dict = {}
    for i in range(1, ws_src.max_row + 1):
        # 处理A-D列 (资产)
        name_a = ws_src[f"A{i}"].value
        if name_a and str(name_a).strip():
            # 先用别名字典翻译，如果找不到，就用原始名称
            name_std = alias_lookup.get(str(name_a).strip(), str(name_a).strip())
            src_dict[name_std] = {
                "期初": ws_src[f"C{i}"].value,
                "期末": ws_src[f"D{i}"].value
            }

        # 处理E-H列 (负债及权益)
        name_e = ws_src[f"E{i}"].value
        if name_e and str(name_e).strip():
            name_std = alias_lookup.get(str(name_e).strip(), str(name_e).strip())
            # 只有当该科目未在资产部分出现时才添加，避免重复
            if name_std not in src_dict:
                 src_dict[name_std] = {
                    "期初": ws_src[f"G{i}"].value,
                    "期末": ws_src[f"H{i}"].value
                }

    logger.debug(f"源Sheet '{sheet_name}' 扫描完成，构建了包含 {len(src_dict)} 个科目的数据字典。")
    
    # --- 3. 将构建好的数据字典转换为标准记录格式 ---
    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    for subject_name, values in src_dict.items():
        # 判断科目类型 (基于我们之前的约定)
        # 注意: 这里的alias_map_df需要重新利用，或在之前步骤构建更复杂结构
        # 为简化，我们暂时只区分我们最关心的几个合计项
        total_subjects = ['资产总计', '负债合计', '净资产合计', '流动资产合计', '非流动资产合计', '流动负债合计', '非流动负债合计']
        subject_type = '合计' if subject_name in total_subjects else '普通'

        record = {
            "来源Sheet": sheet_name,
            "报表类型": "资产负债表",
            "年份": year,
            "项目": subject_name,
            "科目类型": subject_type,
            "期初金额": values["期初"],
            "期末金额": values["期末"]
        }
        records.append(record)
        
    logger.info(f"--- 资产负债表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records