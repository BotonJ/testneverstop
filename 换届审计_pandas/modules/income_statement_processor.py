# /modules/income_statement_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger
from modules.utils import normalize_name

def process_income_statement(ws_src, sheet_name, yewu_line_map, alias_map_df, net_asset_fallback=None):
    """
    【最终版 V4 - 忠于经典代码】
    完整复刻 data_processor-A.py 的健壮逻辑，包括动态行偏移和计算保底。
    """
    logger.info(f"--- 开始处理业务活动表: '{sheet_name}' (使用最终版健壮逻辑) ---")
    
    income_total_aliases = {normalize_name(s) for s in ['收入合计', '一、收 入', '（一）收入合计']}
    expense_total_aliases = {normalize_name(s) for s in ['费用合计', '二、费 用', '（二）费用合计']}
    balance_aliases = {normalize_name(s) for s in ['收支结余', '三、收支结余']}
    net_asset_change_aliases = {normalize_name(s) for s in ['净资产变动额', '五、净资产变动额（若为净资产减少额，以"-"号填列）']}
    
    records = []
    found_items = {}
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    mapping_dict = {normalize_name(item.get("字段名","")): item for item in yewu_line_map} if yewu_line_map else {}

    # --- 步骤 1: 遍历mapping配置，提取所有能提取的数据 ---
    for item_name_clean, row_config in mapping_dict.items():
        if not item_name_clean: continue

        start_coord = row_config.get("源期初坐标")
        end_coord = row_config.get("源期末坐标")
        
        if pd.notna(start_coord) or pd.notna(end_coord):
            try:
                start_val = ws_src[start_coord].value if pd.notna(start_coord) else None
                end_val = ws_src[end_coord].value if pd.notna(end_coord) else None
                
                found_items[item_name_clean] = {"本期": end_val, "上期": start_val}
                
                subject_type = '普通'
                standard_name = item_name_clean
                if item_name_clean in income_total_aliases:
                    standard_name = normalize_name('收入合计')
                    subject_type = '合计'
                elif item_name_clean in expense_total_aliases:
                    standard_name = normalize_name('费用合计')
                    subject_type = '合计'

                records.append({
                    "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
                    "项目": standard_name, "科目类型": subject_type,
                    "本期金额": end_val, "上期金额": start_val
                })
            except Exception as e:
                logger.warning(f"无法提取'{item_name_clean}'的数据，坐标可能无效。错误: {e}")

    # --- 步骤 2: “提取优先，计算保底”逻辑 ---
    found_balance = any(alias in found_items and found_items[alias]["本期"] is not None for alias in balance_aliases)
    if not found_balance:
        income_val = found_items.get(normalize_name('收入合计'), {}).get('本期', 0)
        expense_val = found_items.get(normalize_name('费用合计'), {}).get('本期', 0)
        calculated_balance = (pd.to_numeric(income_val, errors='coerce') or 0) - (pd.to_numeric(expense_val, errors='coerce') or 0)
        records.append({
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": "收支结余", "科目类型": "合计",
            "本期金额": calculated_balance, "上期金额": None
        })
        logger.info(f"自动计算'收支结余'完成，值为: {calculated_balance}")

    found_net_asset_change = any(alias in found_items and found_items[alias]["本期"] is not None for alias in net_asset_change_aliases)
    if not found_net_asset_change and net_asset_fallback:
        start_net = pd.to_numeric(net_asset_fallback.get('期初净资产'), errors='coerce') or 0
        end_net = pd.to_numeric(net_asset_fallback.get('期末净资产'), errors='coerce') or 0
        calculated_change = end_net - start_net
        records.append({
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": "净资产变动额", "科目类型": "合计",
            "本期金额": calculated_change, "上期金额": None
        })
        logger.info(f"自动计算'净资产变动额'完成，值为: {calculated_change}")

    logger.info(f"--- 业务活动表 '{sheet_name}' 处理完成，最终生成 {len(records)} 条记录。---")
    return records