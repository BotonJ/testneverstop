# /modules/income_statement_processor.py
import re
import pandas as pd
from src.utils.logger_config import logger

def process_income_statement(ws_src, sheet_name, yewu_line_map, alias_map_df, net_asset_fallback=None):
    """
    【回溯版 - 忠于原始逻辑】
    采用“提取优先，计算保底”的智能逻辑。
    """
    logger.info(f"--- 开始处理业务活动表: '{sheet_name}' (使用最终版宽容设计) ---")
    
    income_total_aliases = ['收入合计', '一、收 入', '（一）收入合计']
    expense_total_aliases = ['费用合计', '二、费 用', '（二）费用合计']
    balance_aliases = ['收支结余', '三、收支结余']
    net_asset_change_aliases = ['净资产变动额', '五、净资产变动额（若为净资产减少额，以"-"号填列）']

    records = []
    found_items = {}
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    mapping_dict = {}
    if yewu_line_map:
        for item in yewu_line_map:
            if item.get("字段名"):
                mapping_dict[item["字段名"].strip()] = (item.get("源期初坐标"), item.get("源期末坐标"))

    for item_name, coords in mapping_dict.items():
        start_coord, end_coord = coords
        if start_coord and end_coord:
            try:
                start_val = ws_src[start_coord].value
                end_val = ws_src[end_coord].value
                found_items[item_name] = {"本期": end_val, "上期": start_val}
                
                subject_type = '合计' if item_name in income_total_aliases or item_name in expense_total_aliases else '普通'
                standard_name = '收入合计' if item_name in income_total_aliases else ('费用合计' if item_name in expense_total_aliases else item_name)

                records.append({
                    "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
                    "项目": standard_name, "科目类型": subject_type,
                    "本期金额": end_val, "上期金额": start_val
                })
            except Exception:
                logger.warning(f"无法提取'{item_name}'的数据，坐标可能无效: 初'{start_coord}', 末'{end_coord}'")

    # “提取优先，计算保底”逻辑
    found_balance = any(alias in found_items and found_items[alias]["本期"] is not None for alias in balance_aliases)
    if not found_balance:
        income_total = pd.to_numeric(found_items.get('收入合计', {}).get('本期', 0), errors='coerce') or 0
        expense_total = pd.to_numeric(found_items.get('费用合计', {}).get('本期', 0), errors='coerce') or 0
        calculated_balance = income_total - expense_total
        records.append({
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": "收支结余", "科目类型": "合计",
            "本期金额": calculated_balance, "上期金额": None
        })
        logger.info(f"自动计算'收支结余'完成，值为: {calculated_balance}")

    found_net_asset_change = any(alias in found_items and found_items[alias]["本期"] is not None for alias in net_asset_change_aliases)
    if not found_net_asset_change and net_asset_fallback:
        start_net_asset = pd.to_numeric(net_asset_fallback.get('期初净资产'), errors='coerce') or 0
        end_net_asset = pd.to_numeric(net_asset_fallback.get('期末净资产'), errors='coerce') or 0
        calculated_change = end_net_asset - start_net_asset
        records.append({
            "来源Sheet": sheet_name, "报表类型": "业务活动表", "年份": year,
            "项目": "净资产变动额", "科目类型": "合计",
            "本期金额": calculated_change, "上期金额": None
        })
        logger.info(f"自动计算'净资产变动额'完成，值为: {calculated_change}")

    logger.info(f"--- 业务活动表 '{sheet_name}' 处理完成，最终生成 {len(records)} 条记录。---")
    return records