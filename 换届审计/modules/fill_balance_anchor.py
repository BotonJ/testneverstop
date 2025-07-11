from modules.utils import normalize_name

def fill_balance_sheet_by_name(ws_src, ws_tgt, alias_dict, log, skip_list=[]):    
    # ✅ 提取源数据（双列 A-C 和 E-G-H）
    src_dict = {}
    for i in range(1, ws_src.max_row + 1):
        name_a = ws_src[f"A{i}"].value
        if name_a:
            name_std = normalize_name(alias_dict.get(str(name_a).strip(), str(name_a).strip()))
            val_init = ws_src[f"C{i}"].value or ""
            val_final = ws_src[f"D{i}"].value or ""
            src_dict[name_std] = {"期初": val_init, "期末": val_final}

        name_e = ws_src[f"E{i}"].value
        if name_e:
            name_std = normalize_name(alias_dict.get(str(name_e).strip(), str(name_e).strip()))
            val_init = ws_src[f"G{i}"].value or ""
            val_final = ws_src[f"H{i}"].value or ""
            if name_std not in src_dict:
                src_dict[name_std] = {"期初": val_init, "期末": val_final}

    # ✅ 提取模板字段及目标行号
    tgt_dict = {}
    for i in range(1, ws_tgt.max_row + 1):
        name_raw = ws_tgt[f"A{i}"].value
        if name_raw:
            name_std = normalize_name(str(name_raw).strip())
            tgt_dict[name_std] = i

    skip_set = set(normalize_name(n) for n in (skip_list or []))

    for tgt_name, tgt_row in tgt_dict.items():
        if tgt_name in skip_set:
           continue
        if tgt_name in src_dict:            
            try:
                val_init = src_dict[tgt_name]["期初"]
                val_final = src_dict[tgt_name]["期末"]
                ws_tgt[f"B{tgt_row}"].value = val_init
                ws_tgt[f"C{tgt_row}"].value = val_final                         
            except Exception as e:
                pass