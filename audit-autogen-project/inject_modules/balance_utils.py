
from modules.utils import normalize_name

def get_balance_core_data(ws_balance, block_map, alias_dict):
    """
    从资产负债表模板中提取资产总额、负债总额、净资产总额（期初 + 期末），
    使用 mapping["blocks"] 提供的 target_col 初末列 + 动态字段行号识别。
    """
    result = {
        "期初资产总额": 0.0,
        "期末资产总额": 0.0,
        "期初负债总额": 0.0,
        "期末负债总额": 0.0,
        "期初净资产总额": 0.0,
        "期末净资产总额": 0.0,
    }

    # 映射输出字段名 → 模板字段名
    field_map = {
        "资产总额": "资产总计",
        "负债总额": "负债合计",
        "净资产总额": "净资产合计"
    }

    # Step 1：扫描模板 A 列，构造 字段名 → 行号 映射
    row_map = {}
    for i in range(1, ws_balance.max_row + 1):
        val = ws_balance.cell(row=i, column=1).value
        if val:
            std = normalize_name(str(val).strip())
            row_map[std] = i

    # Step 2：定义读取函数
    def safe_get(row, col):
        try:
            val = ws_balance.cell(row=row, column=col).value
            return float(str(val).replace(",", "").strip()) if val not in [None, ""] else 0.0
        except Exception as e:
            print(f"⚠️ 坐标 ({row},{col}) 读取失败: {e}")
            return 0.0

    # Step 3：提取字段值
    for target_field, template_key in field_map.items():
        std_key = normalize_name(template_key)
        alias_keys = [normalize_name(template_key)]
        for k, v in alias_dict.items():
            if normalize_name(k) == std_key:
                alias_keys += [normalize_name(x) for x in v]

        # 优先匹配别名中的任一名称
        row = None
        for name in alias_keys:
            if name in row_map:
                row = row_map[name]
                break

        block = block_map.get(template_key)
        if not block:
            print(f"⚠️ 区块 {template_key} 未找到")
            continue

        col_init = block.get("target_col_initial")
        col_final = block.get("target_col_final")

        if row is None or col_init is None or col_final is None:
            print(f"⚠️ {template_key} 缺少坐标信息（row={row}, col_init={col_init}, col_final={col_final}）")
            continue

        result[f"期初{target_field}"] = safe_get(row, col_init)
        result[f"期末{target_field}"] = safe_get(row, col_final)

    return result
