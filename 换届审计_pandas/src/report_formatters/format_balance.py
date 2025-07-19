# /src/report_formatters/format_balance.py

from modules.utils import normalize_name
import logging

logger = logging.getLogger(__name__)

def format_balance_sheet(ws_src, ws_tgt, alias_dict):    
    """
    【V2 - Bug修复+调试】
    精确填充单个年度的资产负债表。修复了期初/期末错位问题。
    """
    src_dict = {}
    # [修复] 明确从预制件的 B(期初) 和 C(期末) 列读取
    for row in ws_src.iter_rows(min_row=2, values_only=True):
        name_raw = row[0]
        if name_raw:
            # 标准化科目名
            name_std = normalize_name(alias_dict.get(str(name_raw).strip(), str(name_raw).strip()))
            src_dict[name_std] = {"期初": row[1], "期末": row[2]}

    tgt_dict = {}
    for i in range(1, ws_tgt.max_row + 1):
        name_raw = ws_tgt[f"A{i}"].value
        if name_raw:
            tgt_dict[normalize_name(str(name_raw).strip())] = i

    print(f"\n---【调试信息】正在填充 '{ws_tgt.title}' ---")
    for tgt_name, tgt_row in tgt_dict.items():
        if tgt_name in src_dict:
            val_init = src_dict[tgt_name]["期初"]
            val_final = src_dict[tgt_name]["期末"]
            
            # --- [调试打印] ---
            if tgt_name in ["流动资产合计", "资产总计", "负债合计", "净资产合计"]:
                 print(f"  -> 匹配到 '{tgt_name}' (行号 {tgt_row}): 准备写入 期初='{val_init}' 到 B{tgt_row}, 期末='{val_final}' 到 C{tgt_row}")
            
            try:
                # [修复] 确保写入到正确的 B(期初) 和 C(期末) 列
                ws_tgt[f"B{tgt_row}"].value = val_init
                ws_tgt[f"C{tgt_row}"].value = val_final                         
            except Exception as e:
                logger.warning(f"写入 {tgt_name} 到 {ws_tgt.title} 时失败: {e}")
    print("---【调试信息】填充结束 ---\n")