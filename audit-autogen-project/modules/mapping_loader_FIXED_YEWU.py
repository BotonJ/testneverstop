
import pandas as pd
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string

def load_mapping_file(mapping_file_path):
    xls = pd.read_excel(mapping_file_path, sheet_name=None)

    # === 提取资产负债表区块 ===
    blocks_df = xls.get("资产负债表区块", pd.DataFrame())
    blocks = {}
    for _, row in blocks_df.iterrows():
        name = str(row.get("区块名称")).strip()
        block = {
            "start_cell": str(row.get("起始单元格")).strip(),
            "end_cell": str(row.get("终止单元格")).strip(),
            "sorce_col_initial": col_letter_to_index(row.get("源期初列")),
            "sorce_col_final": col_letter_to_index(row.get("源期末列")),
            "target_col_initial": col_letter_to_index(row.get("目标期初列")),
            "target_col_final": col_letter_to_index(row.get("目标期末列")),
            "target_start_cell": str(row.get("目标起始单元格")).strip(),
            "target_end_cell": str(row.get("目标终止单元格")).strip(),
            "skip_rows": [],
        }
        try:
            _, row_number = coordinate_from_string(block["target_start_cell"])
            block["target_row"] = int(row_number)
        except Exception:
            block["target_row"] = None
        blocks[name] = block

    # === 提取科目等价映射 ===
    alias_map_df = xls.get("科目等价映射", pd.DataFrame())
    alias_dict = {}
    for _, row in alias_map_df.iterrows():
        std = str(row.get("标准科目名", "")).strip()
        if not std:
            continue
        for col in row.index:
            if col.startswith("等价科目名"):
                alias = str(row[col]).strip()
                if alias:
                    alias_dict[alias] = std

    # === 提取业务活动表逐行 ===
    yewu_df = xls.get("业务活动表逐行", pd.DataFrame())
    yewu_mapping = []
    for _, row in yewu_df.iterrows():
        yewu_mapping.append({
            "字段名": str(row.get("字段名", "")).strip(),
            "源期初坐标": str(row.get("源期初坐标", "")).strip(),
            "源期末坐标": str(row.get("源期末坐标", "")).strip(),
            "目标期初坐标": str(row.get("目标期初坐标", "")).strip(),
            "目标期末坐标": str(row.get("目标期末坐标", "")).strip(),
            "是否计算": str(row.get("是否计算", "")).strip(),
        })

    return {
        "blocks": blocks,
        "subject_alias_map": alias_dict,
        "yewu_mapping": yewu_mapping
    }

def col_letter_to_index(letter):
    if pd.isna(letter):
        return None
    try:
        return column_index_from_string(str(letter).strip())
    except Exception:
        print(f"⚠️ 无法转换列标 '{letter}'")
        return None
