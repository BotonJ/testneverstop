# /src/report_formatters/format_yewu.py

import pandas as pd
import logging
from modules.utils import normalize_name

logger = logging.getLogger(__name__)

def format_yewu_sheet(ws_src, ws_tgt, yewu_line_map, prev_ws=None, net_asset_change=None, is_first_audit_year=False):
    """
   【V7 - 最终版】
    - 适配标准化的业务活动表预制件，该预制件现在包含期初和期末两列。
    - 明确区分审计第一年和后续年份的数据来源和处理逻辑。
    - 修复了审计第一年期初有值但未写入的BUG。
    """
    is_map_valid = yewu_line_map is not None and not yewu_line_map.empty
    if not is_map_valid: return
    config_list = yewu_line_map.to_dict('records')

    # 1. 【数据准备】将本年度的预制件(ws_src)数据读入一个字典，方便快速查找
    src_data_map = {}
    # 预制件从第2行开始是数据，A列是项目，B列是期初，C列是期末
    for row in ws_src.iter_rows(min_row=2, max_col=3, values_only=True):
        item_name = row[0]
        if item_name:
            src_data_map[normalize_name(item_name)] = {"期初": row[1], "期末": row[2]}

    # 2. 【填充期初列】根据是否为审计第一年，选择不同逻辑
    print(f"\n---【调试信息】正在填充 '{ws_tgt.title}' 的期初列 ---")
    if is_first_audit_year:
        print("  -> 当前为审计第一年，将从本年预制件的'期初'列获取数据并进行有效性检查...")
        # 预读所有期初值，判断是否“实质为空”
        qichu_values = [src_data_map.get(normalize_name(item.get("字段名")), {}).get("期初") for item in config_list]
        
        # any() 对数字0也视为False, 所以要处理一下
        has_real_values = any(val for val in qichu_values if val is not None and float(val) != 0)

        if has_real_values:
            print("  -> 期初数据有效（存在非零值），开始逐项写入...")
            for item in config_list:
                tgt_qichu_coord = item.get("目标期初坐标")
                field_name_norm = normalize_name(item.get("字段名"))
                if tgt_qichu_coord and field_name_norm in src_data_map:
                    ws_tgt[tgt_qichu_coord].value = src_data_map[field_name_norm].get("期初")
        else:
            print("  -> [警告] 审计第一年的期初数据所有科目值均为0或None，判定为“实质为空”。将跳过填充期初列。")
    
    elif prev_ws:
        print("  -> 当前非审计第一年，将严格继承上年期末数据...")
        for item in config_list:
            tgt_qichu_coord = item.get("目标期初坐标")
            src_qimo_coord = item.get("目标期末坐标")
            if tgt_qichu_coord and src_qimo_coord:
                try: ws_tgt[tgt_qichu_coord].value = prev_ws[src_qimo_coord].value
                except KeyError: pass
 
     #3. 【填充期末列】和【4. 执行计算】
    print(f"---【调试信息】正在填充 '{ws_tgt.title}' 的期末列及计算项 ---")
    for item in config_list:
        field_norm = normalize_name(item.get("字段名"))
        is_calc = str(item.get("是否计算", "")).strip() == "是"

        # --- [核心修正] ---
        # 将目标坐标的定义移到循环的顶部，使其对'if is_calc'和'else'块都可用
        tgt_qimo_coord = item.get("目标期末坐标")

        # 增加一个安全检查，如果配置行没有目标坐标，则跳过
        if not tgt_qimo_coord:
            continue

        if is_calc:
            # 计算“收支结余”
            if "收支结余" in field_norm:
                try:
                    income_tgt_coord = next(i["目标期末坐标"] for i in config_list if "收入合计" in normalize_name(i["字段名"]))
                    expense_tgt_coord = next(i["目标期末坐标"] for i in config_list if "费用合计" in normalize_name(i["字段名"]))
                    income = float(ws_tgt[income_tgt_coord].value or 0)
                    expense = float(ws_tgt[expense_tgt_coord].value or 0)
                    balance = income - expense
                    print(f"  -> 计算'收支结余': 收入({income_tgt_coord})={income}, 费用({expense_tgt_coord})={expense}, 结余={balance}")
                    # 使用在循环顶部定义的坐标进行写入
                    ws_tgt[tgt_qimo_coord].value = balance
                except Exception as e:
                    print(f"  -> [错误] 计算'收支结余'失败: {e}")

            # 计算“净资产变动额”
            elif "净资产变动额" in field_norm:
                if net_asset_change is not None:
                    print(f"  -> 写入'净资产变动额': 使用传入的值 {net_asset_change} 写入到 {tgt_qimo_coord}")
                    # 使用在循环顶部定义的坐标进行写入
                    ws_tgt[tgt_qimo_coord].value = net_asset_change
                else:
                    print(f"  -> [警告] '净资产变动额' 的计算值 (net_asset_change) 为None，跳过写入。")
        
        else: # 如果不是计算项，则直接填充
            if field_norm in src_data_map:
                ws_tgt[tgt_qimo_coord].value = src_data_map[field_norm].get("期末")

    print("---【调试信息】填充结束 ---\n")