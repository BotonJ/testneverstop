import os
import openpyxl
import pandas as pd

os.chdir(os.path.dirname(os.path.abspath(__file__)))

#统一数字格式
def apply_number_format(cell):
    cell.number_format = '#,##0.00'
# 1. 读取 inj1/inj2/inj3 Sheet 的配置参数和区块定义
def get_mapping_conf_and_df(mapping_file, sheet_name):
    df = pd.read_excel(mapping_file, sheet_name=sheet_name, header=None)
    conf = {}
    data_start = 0
    for i, row in df.iterrows():
        values = [str(cell).strip() for cell in row if pd.notna(cell)]
        if any(k in values for k in ["起始行", "来源字段"]):
            data_start = i
            break
        if pd.notna(row[0]) and pd.notna(row[1]):
            conf[str(row[0]).strip()] = str(row[1]).strip()
    df_data = pd.read_excel(mapping_file, sheet_name=sheet_name, header=data_start)
    return conf, df_data

# 2. 表1注入函数（固定字段注入）
def inject_table1(wb_src, ws_tgt, conf, df_map):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        src_field = str(row["来源字段"]).strip()
        tgt_init_cell = str(row["目标单元格（期初）"]).strip()
        tgt_final_cell = str(row["目标单元格（期末）"]).strip()
        var_cell = str(row.get("变动单元格", "")).strip()
        var_formula = str(row.get("变动公式", "")).strip()

        val_init, val_final = None, None
        for i in range(1, 100):
            name = ws_src_init[f"A{i}"].value
            if name and src_field in str(name):
                val_init = ws_src_init[f"B{i}"].value
                break
        for i in range(1, 100):
            name = ws_src_final[f"A{i}"].value
            if name and src_field in str(name):
                val_final = ws_src_final[f"C{i}"].value
                break

        ws_tgt[tgt_init_cell] = val_init
        apply_number_format(ws_tgt[tgt_init_cell])
        ws_tgt[tgt_final_cell] = val_final
        apply_number_format(ws_tgt[tgt_final_cell])
        if var_cell and var_formula:
            ws_tgt[var_cell] = var_formula

# 3. 表2/4注入函数（区块遍历）
def inject_table2(wb_src, ws_tgt, conf, df_map):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        start_row = int(row["起始行"])
        end_row = int(row["终止行"])
        src_col_init = str(row["来源列（期初）"]).strip()
        src_col_final = str(row["来源列（期末）"]).strip()
        tgt_row = int(row["目标起始单元格"][1:])
        tgt_col_prefix = str(row["目标起始单元格"][0]).strip()
        var_col = str(row.get("变动单元格起始", "")).strip()
        var_formula = str(row.get("变动公式", "")).strip()
        skip_strs = [s.strip() for s in str(row.get("跳过行", "")).split(",") if s.strip()]
        skip_zero = str(row.get("是否跳过均为0", "")).strip() == "是"
        out_row = tgt_row
        write_count = 0
        for r in range(start_row, end_row + 1):
            subject = ws_src_init[f"A{r}"].value
            if not subject or any(skip in subject for skip in skip_strs):
                continue
            val_init = ws_src_init[f"{src_col_init}{r}"].value
            val_final = ws_src_final[f"{src_col_final}{r}"].value
            if skip_zero and ((not val_init or val_init == 0) and (not val_final or val_final == 0)):
                continue
            ws_tgt[f"{tgt_col_prefix}{out_row}"] = subject
            ws_tgt[f"{chr(ord(tgt_col_prefix)+1)}{out_row}"] = val_init if val_init is not None else ""
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+1)}{out_row}"])
            ws_tgt[f"{chr(ord(tgt_col_prefix)+2)}{out_row}"] = val_final if val_final is not None else ""
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+2)}{out_row}"])
            ws_tgt[f"{chr(ord(tgt_col_prefix)+3)}{out_row}"] = f"={chr(ord(tgt_col_prefix)+2)}{out_row}-{chr(ord(tgt_col_prefix)+1)}{out_row}"
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+3)}{out_row}"])
            if var_formula and var_col:
                cell = var_col[0] + str(out_row)
                ws_tgt[cell] = var_formula.replace("11", str(out_row))
            write_count += 1
            out_row += 1
              
        if write_count > 0:
            sum_label = str(row.get("合计行名称", "")).strip() or f"{row.get('区块名称', '')}合计"
            ws_tgt[f"{tgt_col_prefix}{out_row}"] = sum_label           
            ws_tgt[f"{chr(ord(tgt_col_prefix)+1)}{out_row}"] = f"=SUM({chr(ord(tgt_col_prefix)+1)}{tgt_row}:{chr(ord(tgt_col_prefix)+1)}{out_row - 1})"
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+1)}{out_row}"])
            ws_tgt[f"{chr(ord(tgt_col_prefix)+2)}{out_row}"] = f"=SUM({chr(ord(tgt_col_prefix)+2)}{tgt_row}:{chr(ord(tgt_col_prefix)+2)}{out_row - 1})"
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+2)}{out_row}"])
            ws_tgt[f"{chr(ord(tgt_col_prefix)+3)}{out_row}"] = f"=SUM({chr(ord(tgt_col_prefix)+3)}{tgt_row}:{chr(ord(tgt_col_prefix)+3)}{out_row - 1})"
            apply_number_format(ws_tgt[f"{chr(ord(tgt_col_prefix)+3)}{out_row}"])   
                  
# 4. 表3注入函数

def inject_table3(wb_src, ws_tgt, conf, df_map):
    start_sheet = conf.get("start_sheet")
    end_sheet = conf.get("end_sheet")
    ws_src_init = wb_src[start_sheet]
    ws_src_final = wb_src[end_sheet]

    for idx, row in df_map.iterrows():
        if pd.isna(row.get("来源字段")):  # 合计行留给后面
            continue
        src_init_cell = str(row.get("来源单元格（期初）", "")).strip()
        src_final_cell = str(row.get("来源单元格（期末）", "")).strip()
        tgt_init_cell = str(row.get("目标单元格（期初）", "")).strip()
        tgt_final_cell = str(row.get("目标单元格（期末）", "")).strip()
        var_formula = str(row.get("变动公式", "")).strip()
        add_cell = str(row.get("增加单元格", "")).strip()
        sub_cell = str(row.get("减少单元格", "")).strip()

        val_init = ws_src_init[src_init_cell].value if src_init_cell else 0
        val_final = ws_src_final[src_final_cell].value if src_final_cell else 0

        if tgt_init_cell:
            ws_tgt[tgt_init_cell] = val_init
            apply_number_format(ws_tgt[tgt_init_cell])
        if tgt_final_cell:
            ws_tgt[tgt_final_cell] = val_final
            apply_number_format(ws_tgt[tgt_final_cell])

        try:
            num_init = float(val_init) if val_init not in [None, ""] else 0
            num_final = float(val_final) if val_final not in [None, ""] else 0
            diff = num_final - num_init
            if diff > 0 and add_cell:
                ws_tgt[add_cell] = diff
            elif diff <= 0 and sub_cell:
                ws_tgt[sub_cell] = abs(diff)
        except Exception as e:
            print(f"⚠️ 表3净资产差值写入失败：{e}")

    for idx, row in df_map.iterrows():
        if pd.notna(row.get("来源字段")):
            continue
        cell = str(row.get("变动单元格", "")).strip()
        formula = str(row.get("变动公式", "")).strip()
        if cell and formula:
            ws_tgt[cell] = formula
            apply_number_format(ws_tgt[cell])
            print(f"✅ 表3 合计列公式写入 {cell} = {formula}")

# 5. 通用合计公式 Sheet 注入

def inject_formula_sheet(ws_tgt, mapping_file):
    try:
        df = pd.read_excel(mapping_file, sheet_name="合计公式配置")
        for _, row in df.iterrows():
            cell = str(row.get("变动单元格", "")).strip()
            formula = str(row.get("变动公式", "")).strip()
            if cell and formula:
                ws_tgt[cell] = formula
                apply_number_format(ws_tgt[cell])
                print(f"✅ 合计公式配置写入 {cell} = {formula}")
    except Exception as e:
        print("⚠️ 未找到或读取合计公式配置失败：", e)

# 主流程
mapping_file = "mapping_file.xlsx"
source_file = "source_file.xlsx"
template_file = "template_file.xlsx"
output_file = "分析模板_自动填报.xlsx"

wb_src = openpyxl.load_workbook(source_file, data_only=True)
wb_tgt = openpyxl.load_workbook(template_file)
ws_tgt = wb_tgt.active

conf1, df1 = get_mapping_conf_and_df(mapping_file, "inj1")
conf2, df2 = get_mapping_conf_and_df(mapping_file, "inj2")
conf3, df3 = get_mapping_conf_and_df(mapping_file, "inj3")

inject_table1(wb_src, ws_tgt, conf1, df1)
inject_table2(wb_src, ws_tgt, conf2, df2)
inject_table3(wb_src, ws_tgt, conf3, df3)
inject_formula_sheet(ws_tgt, mapping_file)

wb_tgt.save(output_file)
print(f"✅ 分析模板自动填报完成：{output_file}")
