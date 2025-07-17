# /modules/income_statement_processor.py

import pandas as pd
import re
from src.utils.logger_config import logger
from modules.utils import normalize_name, _get_value_from_coord

def process_income_statement(sheet, sheet_name, alias_map_df, is_map_df, is_summary_config_df):
    """
    【V2 - Bug修复】
    - 修正了合计科目类型(subject_type)被错误覆盖为'普通'的BUG。
    - 使用与资产负债表一致的、更健壮的科目类型判断逻辑。
    """
    logger.info(f"--- 开始处理业务活动表: '{sheet_name}' (使用V2'Bug修复'逻辑) ---")
    
    if not is_map_df:
        logger.warning(f"'{sheet_name}': '业务活动表逐行'配置为空，跳过处理。")
        return []

    # 从 '业务活动表汇总注入配置' 创建类型查找字典（收入/费用）
    type_lookup = {}
    if is_summary_config_df is not None and not is_summary_config_df.empty:
        for _, row in is_summary_config_df.iterrows():
            clean_name = normalize_name(row['科目名称'])
            if clean_name:
                type_lookup[clean_name] = row['类型']

    # --- [修复] 创建两个独立的查找结构，同 balance_sheet_processor.py ---
    alias_lookup = {}
    total_items_set = set() # 专门存放所有类型为'合计'的标准科目名
    if alias_map_df is not None and not alias_map_df.empty:
        for _, row in alias_map_df.iterrows():
            standard_clean = normalize_name(row['标准科目名'])
            if not standard_clean: continue
            
            is_total = '科目类型' in row and str(row['科目类型']).strip() == '合计'
            if is_total:
                total_items_set.add(standard_clean)
            
            alias_lookup[standard_clean] = standard_clean
            for col in alias_map_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    aliases = [normalize_name(alias) for alias in str(row[col]).split(',')]
                    for alias in aliases:
                        if alias: alias_lookup[alias] = standard_clean

    records = []
    year_match = re.search(r'(\d{4})', sheet_name)
    year = int(year_match.group(1)) if year_match else "未知"

    for yewu_line in is_map_df:
        subject_name_raw = yewu_line.get('字段名')
        if not subject_name_raw: continue

        subject_name_clean = normalize_name(subject_name_raw)
        
        # --- [修复] 两步式判断，确保类型正确 ---
        # 第1步：确定标准名称
        standard_name = alias_lookup.get(subject_name_clean, subject_name_raw) # 保留原始名以防万一
        
        # 第2步：根据标准名称，使用 total_items_set 判断其类型
        subject_type = '合计' if normalize_name(standard_name) in total_items_set else '普通'
        
        # 使用 type_lookup 确定科目分类（收入/费用）
        item_type = type_lookup.get(normalize_name(standard_name), '')

        start_val = _get_value_from_coord(sheet, yewu_line.get('源期初坐标'))
        end_val = _get_value_from_coord(sheet, yewu_line.get('源期末坐标'))

        record = {
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": standard_name, "所属区块": item_type, "科目类型": subject_type, # 使用修复后的 subject_type
            "类型": item_type, "期初金额": start_val, "期末金额": end_val
        }
        records.append(record)

    # 自动计算“净资产变动额”
    try:
        income_total = sum(r['期末金额'] for r in records if r.get('类型') == '收入' and pd.notna(r.get('期末金额')))
        expense_total = sum(r['期末金额'] for r in records if r.get('类型') == '费用' and pd.notna(r.get('期末金额')))
        net_change = income_total - expense_total
        logger.info(f"自动计算'净资产变动额'完成，值为: {net_change}")
        records.append({
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": "净资产变动额", "所属区块": "结余", "科目类型": "合计", "类型": "结余",
            "期初金额": None, "期末金额": net_change
        })
    except Exception as e:
        logger.warning(f"自动计算'净资产变动额'失败: {e}")

    logger.info(f"--- 业务活动表 '{sheet_name}' 处理完成，最终生成 {len(records)} 条记录。---")
    return records