# /src/report_formatters/format_balance.py

from modules.utils import normalize_name

def format_balance_sheet(ws_src, ws_tgt, alias_dict):    
    """
    精确填充单个年度的资产负债表。
    :param ws_src: 包含该年度干净数据的源Sheet (A/C/D, E/G/H格式)
    :param ws_tgt: 从t.xlsx复制过来的目标Sheet
    :param alias_dict: 别名->标准名映射
    """
    src_dict = {}
    for i in range(1, ws_src.max_row + 1):
        # 左侧
        name_a = ws_src[f"A{i}"].value
        if name_a:
            name_std = normalize_name(alias_dict.get(str(name_a).strip(), str(name_a).strip()))
            src_dict[name_std] = {"期初": ws_src[f"C{i}"].value, "期末": ws_src[f"D{i}"].value}
        # 右侧
        name_e = ws_src[f"E{i}"].value
        if name_e:
            name_std = normalize_name(alias_dict.get(str(name_e).strip(), str(name_e).strip()))
            if name_std not in src_dict:
                src_dict[name_std] = {"期初": ws_src[f"G{i}"].value, "期末": ws_src[f"H{i}"].value}

    tgt_dict = {}
    for i in range(1, ws_tgt.max_row + 1):
        name_raw = ws_tgt[f"A{i}"].value
        if name_raw:
            tgt_dict[normalize_name(str(name_raw).strip())] = i

    for tgt_name, tgt_row in tgt_dict.items():
        if tgt_name in src_dict:            
            try:
                ws_tgt[f"B{tgt_row}"].value = src_dict[tgt_name]["期初"]
                ws_tgt[f"C{tgt_row}"].value = src_dict[tgt_name]["期末"]                         
            except Exception:
                pass