
import pandas as pd
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def fmt(val):
    try:
        return f"{float(val):,.2f}"
    except:
        return "0.00"

def build_context_and_values(activity_df, balance_df, map_df):
    """
    根据三张数据表生成：
    1. context 字典（供 docxtpl 渲染使用）
    2. values 字典（原始数据，供附注表格/描述使用）
    """
    map_df['column'] = map_df['column'].fillna('本年累计数_合计').astype(str).str.strip()
    map_df['project_name'] = map_df['project_name'].astype(str).str.strip()
    map_df['source_sheet'] = map_df['source_sheet'].astype(str).str.strip()

    values = {}
    for _, row in map_df.iterrows():
        sheet = row['source_sheet']
        item  = row['project_name']
        col   = row['column']
        df = activity_df if sheet == "业务活动表" else balance_df if sheet == "资产负债表" else None
        if df is not None:
            m = df[df.iloc[:,0] == item]
            val = m.iloc[0][col] if (not m.empty and col in df.columns) else None
        else:
            val = None
        values[(sheet, item)] = val

    context = {}
    for _, row in map_df.iterrows():
        key = row['context_key']
        context[key] = fmt(values.get((row['source_sheet'], row['project_name']), 0))

    return context, values
