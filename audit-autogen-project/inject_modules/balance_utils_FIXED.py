def get_balance_core_data(ws_balance, block_map, alias_dict):
    result = {
        "期初资产总额": 0.0,
        "期末资产总额": 0.0,
        "期初负债总额": 0.0,
        "期末负债总额": 0.0,
        "期初净资产总额": 0.0,
        "期末净资产总额": 0.0,
    }

    field_map = {
        "资产总额": "资产总计",
        "负债总额": "负债合计",
        "净资产总额": "净资产合计"
    }

    def safe_get(row, col):
        try:
            val = ws_balance.cell(row=row, column=col).value
            if isinstance(val, str) and val.strip().startswith("="):
                print(f"⚠️ 单元格 ({row},{col}) 为公式 → 返回 0")
                return 0.0
            return float(str(val).replace(",", "").strip()) if val not in [None, ""] else 0.0
        except Exception as e:
            print(f"⚠️ 坐标 ({row},{col}) 读取失败: {e}")
            return 0.0

    for output_field, mapped_name in field_map.items():
        std_key = alias_dict.get(mapped_name, mapped_name)
        block = block_map.get(std_key)
        if not block:
            print(f"⚠️ 未匹配字段 {output_field}，block 缺失")
            continue

        row = block.get("target_row")
        col_qichu = block.get("target_col_initial")
        col_qimo = block.get("target_col_final")
        if not row or not col_qichu or not col_qimo:
            print(f"⚠️ 未匹配字段 {output_field}，坐标不全")
            continue

        result[f"期初{output_field}"] = safe_get(row, col_qichu)
        result[f"期末{output_field}"] = safe_get(row, col_qimo)

    return result
