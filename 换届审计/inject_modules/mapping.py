import pandas as pd

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
