from modules.log_utils import log_write
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

def safe_read(ws, cell_ref):
    try:
        col_letter, row = coordinate_from_string(cell_ref)
        col = column_index_from_string(col_letter)
        return ws.cell(row=row, column=col).value
    except Exception:
        return "-"

def fill_yewu_by_mapping(ws_src, ws_tgt,yewu_line_map,prev_ws=None, net_asset_fallback=None, log=None):
    if log is not None:
        log.append("✅ fill_yewu_by_mapping 已启动")       
    for item in yewu_line_map:
        field = item.get("字段名")
        src_initial = item.get("源期初坐标")
        src_final = item.get("源期末坐标")
        tgt_initial = item.get("目标期初坐标")
        tgt_final = item.get("目标期末坐标")
        is_calc = str(item.get("是否计算", "")).strip() == "是"    
        # 归档前补: 连续前一年的期末值
        if prev_ws and tgt_initial and tgt_final:
            try:
                prev_val = prev_ws[tgt_final].value
                ws_tgt[tgt_initial].value = prev_val                
            except Exception as e:
                print(f"⚠️ 行列前年期末补充失败: {field}, {e}")

        # 🧶 收支结余
        if is_calc:
            if "收支结余" in str(field):
                try:
                    income_coord = next((i["目标期末坐标"] for i in yewu_line_map if str(i["字段名"]).strip() == "收 入 合 计"), None)
                    expense_coord = next((i["目标期末坐标"] for i in yewu_line_map if str(i["字段名"]).strip() == "费 用 合 计"), None)
                    income = ws_tgt[income_coord].value if income_coord else None
                    expense = ws_tgt[expense_coord].value if expense_coord else None
                    income = float(income) if income not in (None, "") else 0
                    expense = float(expense) if expense not in (None, "") else 0
                    result = round(income - expense, 2)
                    ws_tgt[tgt_final].value = result                    
                    if log:
                        log_write(log, "success", field, f"收支结余计算: {income} - {expense} = {result} → 写入: {tgt_final}")
                except Exception as e:
                    print(f"❌ 收支结余计算失败: {e}")
            elif "净资产变动额" in str(field) and net_asset_fallback:
                try:
                    val_initial = net_asset_fallback.get("期初", 0)
                    val_final = net_asset_fallback.get("期末", 0)
                    result = round(val_final - val_initial, 2)
                    ws_tgt[tgt_final].value = result
                    if log:
                        log_write(log, "success", field, f"使用资产负债表 fallback: {val_final} - {val_initial} = {result} → 写入: {tgt_final}")
                except Exception as e:
                    if log:
                        log_write(log, "error", field, f"净资产 fallback 计算失败: {e}")
                continue

        # 正常期初值写入
        if src_initial and tgt_initial:
            try:
                ws_tgt[tgt_initial].value = ws_src[src_initial].value
            except Exception as e:
                print(f"⚠️ 期初写入失败: {field}, {e}")

        # 正常期末值写入
        if src_final and tgt_final:
            try:
                ws_tgt[tgt_final].value = ws_src[src_final].value
            except Exception as e:
                print(f"⚠️ 期末写入失败: {field}, {e}")
            if field in ["收 入 合 计", "费 用 合 计", "收支结余"]:
                try:
                    val = ws_tgt[tgt_final].value
                    log_write(log, "success", f"{ws_tgt.title} {field}", f"期末={val}")
                except Exception as e:
                    log_write(log, "error", f"{ws_tgt.title} {field}", f"读取失败: {e}")

        # 无误情况下输出日志
        if log:
            val_i = safe_read(ws_src, src_initial)
            val_f = safe_read(ws_src, src_final)
            log_write(log, "success", field, f"期初={val_i}, 期末={val_f} → {tgt_initial}, {tgt_final}")
