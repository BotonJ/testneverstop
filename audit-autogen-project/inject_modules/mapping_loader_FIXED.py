from openpyxl.utils.cell import coordinate_from_string
import pandas as pd

def load_mapping_file(mapping_file_path):
    xls = pd.read_excel(mapping_file_path, sheet_name=None)
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
        # ✅ 提取 target_row
        try:
            _, row_number = coordinate_from_string(block["target_start_cell"])
            block["target_row"] = int(row_number)
        except Exception:
            block["target_row"] = None

        blocks[name] = block

    alias_map_df = xls.get("科目等价映射", pd.DataFrame())
    alias_dict = {}
    for _, row in alias_map_df.iterrows():
        std = str(row["标准科目名"]).strip()
        alias = str(row["等价科目名"]).strip()
        if alias:
            alias_dict[alias] = std

    return {"blocks": blocks, "subject_alias_map": alias_dict}

def col_letter_to_index(letter):
    if pd.isna(letter):
        return None
    letter = str(letter).strip().upper()
    return ord(letter) - ord('A') + 1
