# /src/report_formatters/format_yewu.py

def format_yewu_sheet(ws_src, ws_tgt, yewu_line_map, prev_ws=None, net_asset_change=None):
    """
    【V2 - Bug修复】
    精确填充单个年度的业务活动表，处理数据衔接和内部计算。
    修复了因对DataFrame进行模糊布尔判断导致的ValueError。
    """
    # --- [修复] 使用 .empty 来判断DataFrame是否为空 ---
    is_map_valid = yewu_line_map is not None and not yewu_line_map.empty
    
    # 填充上一年的期末值到本年的期初
    if prev_ws and is_map_valid:
        for _, item in yewu_line_map.iterrows(): # 遍历DataFrame的行
            tgt_initial = item.get("目标期初坐标")
            tgt_final = item.get("目标期末坐标")
            if tgt_initial and tgt_final and prev_ws[tgt_final].value is not None:
                try:
                    ws_tgt[tgt_initial].value = prev_ws[tgt_final].value
                except Exception:
                    pass

    # 填充本年数据和计算项
    if is_map_valid:
        # 先把所有配置项转成一个list of dicts，提高后续查找效率
        config_list = yewu_line_map.to_dict('records')
        
        for item in config_list:
            field = item.get("字段名", "")
            src_coord = item.get("源期末坐标")
            tgt_coord = item.get("目标期末坐标")
            is_calc = str(item.get("是否计算", "")).strip() == "是"

            if is_calc:
                # 处理“收支结余”
                if "收支结余" in field:
                    try:
                        income_coord = next(i["目标期末坐标"] for i in config_list if "收 入 合 计" in i["字段名"])
                        expense_coord = next(i["目标期末坐标"] for i in config_list if "费 用 合 计" in i["字段名"])
                        income = float(ws_tgt[income_coord].value or 0)
                        expense = float(ws_tgt[expense_coord].value or 0)
                        ws_tgt[tgt_coord].value = income - expense
                    except (StopIteration, TypeError, KeyError):
                        pass
                # 处理“净资产变动额”
                elif "净资产变动额" in field and net_asset_change is not None:
                    ws_tgt[tgt_coord].value = net_asset_change
            elif src_coord and tgt_coord:
                # 正常从源(预制件)填充数据
                try:
                    ws_tgt[tgt_coord].value = ws_src[src_coord].value
                except KeyError:
                    pass