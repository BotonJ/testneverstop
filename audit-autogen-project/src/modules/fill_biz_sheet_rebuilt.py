
import pandas as pd
from openpyxl.utils import get_column_letter
from modules.cell_utils import get_value_by_coord, write_value_by_coord

def fill_biz_sheet(ws_src, ws_tgt, ws_tgt_name, all_ws_biz, biz_map_df, alias_dict, log):
    for idx, row in biz_map_df.iterrows():
        field = row["字段名"]
        src_coord = row.get("源单元格", "")
        tgt_coord = row.get("目标单元格", "")
        compute = row.get("是否计算", "").strip() == "是"

        # 支持自定义计算字段
        if compute and field == "收支结余":
            try:
                income = get_value_by_coord(ws_src, alias_dict.get("收入合计", "H35"))
                expense = get_value_by_coord(ws_src, alias_dict.get("支出合计", "H54"))
                result = income - expense
                write_value_by_coord(ws_tgt, tgt_coord, result)
                log.append(f"✅ 自动计算字段：{field} = {income} - {expense} → {result}")
            except Exception as e:
                log.append(f"⚠️ 无法计算字段 {field}：{e}")
            continue

        # 正常读取流程（增强容错）
        try:
            value = get_value_by_coord(ws_src, src_coord)
            write_value_by_coord(ws_tgt, tgt_coord, value)
            log.append(f"✅ 字段：{field} 写入成功 → {value}")
        except Exception as e:
            log.append(f"⚠️ 字段 {field} 写入失败：{e}")
