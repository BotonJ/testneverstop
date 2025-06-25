from openpyxl.utils import coordinate_to_tuple, column_index_from_string
import re

def render_header(wb_tgt, sheet_name, year, header_meta, unit_name=None):
    ws = wb_tgt[sheet_name]
    sheet_key = "资产负债表" if "资产负债表" in sheet_name else "业务活动表"  
    for field_name, meta in header_meta.items():
        typ = str(meta.get("type", "")).strip()
        rule = meta.get("rule", "").strip()
        cells = meta.get("target_cells", {}).get(sheet_key, [])        
        if not cells:
            continue
        value = None

        # ✅ 写入单位名称（优先使用 rule）
        if "单位" in typ:
            if rule:
                value = rule
            elif unit_name:
                value = f"编制单位：{unit_name}"
            else:                
                continue

        # ✅ 写入期初 / 期末
        elif "期初" in typ or "期末" in typ:
            is_balance = "资产负债表" in sheet_name
            is_biz = "业务活动表" in sheet_name
            # 提取审计末年与末月（如 rule = "2021年7月-2025年X月"）
            audit_end = None
            audit_end_month = None
            match = re.match(r"(\d{4})年.*-(\d{4})年(\d{1,2})月", rule)
            if match:
                audit_end = int(match.group(2))
                audit_end_month = int(match.group(3))

            if is_balance:
                if "期初" in typ:
                    value = f"{year - 1}年12月31日"
                elif "期末" in typ:
                    value = f"{year}年12月31日"
            elif is_biz:
                if "期初" in typ:
                    value = f"{year - 1}年累计数"
                elif "期末" in typ:
                    if audit_end and year == audit_end:
                        month_str = f"1-{audit_end_month}月" if audit_end_month else "1-3月"
                        value = f"{year}年{month_str}累计数"
                    else:
                        value = f"{year}年累计数"                        

        if not value:           
            continue

        for r, c in cells:
            if r and c:
                ws.cell(row=r, column=c, value=value)
               


