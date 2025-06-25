import pandas as pd

mapping_file = "tt.xlsx"

# 要检测的所有注入sheet名称
sheet_names = ["inj1", "inj2", "inj3"]

for sheet in sheet_names:
    try:
        df = pd.read_excel(mapping_file, sheet_name=sheet)
        print(f"\n[{sheet}] 字段名：\n", df.columns.tolist())
    except Exception as e:
        print(f"\n[{sheet}] 读取出错：{e}")
