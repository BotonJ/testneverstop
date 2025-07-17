# /src/report_formatters/format_biz_summary.py

import pandas as pd
import logging
import re

def create_and_inject_biz_summary(prebuilt_wb, wb_to_fill, mapping_configs):
    """
    汇总所有年份的业务活动数据，生成格式化的汇总表，并注入。
    """
    logging.info("  -> 生成并注入 '收入汇总' 和 '支出汇总' Sheet...")
    try:
        # 1. 读取科目配置
        df_subjects = mapping_configs['业务活动表汇总注入配置']
        income_subjects = df_subjects[df_subjects['类型'] == '收入']['科目名称'].tolist()
        expense_subjects = df_subjects[df_subjects['类型'] == '支出']['科目名称'].tolist()
        
        # 2. 从预制件中收集数据
        all_data = []
        for sheet_name in prebuilt_wb.sheetnames:
            if '业务活动表' in sheet_name:
                year_str = re.search(r'\d{4}', sheet_name).group(0)
                ws = prebuilt_wb[sheet_name]
                for row in ws.iter_rows(min_row=2):
                    subject = row[0].value
                    amount = row[3].value # D列是本期金额
                    if subject and amount is not None:
                        all_data.append([year_str, subject, amount])
        
        if not all_data: return
        
        full_df = pd.DataFrame(all_data, columns=['年份', '科目', '金额'])
        full_df['金额'] = pd.to_numeric(full_df['金额'], errors='coerce').fillna(0)

        # 3. 创建并注入收入/支出汇总表
        _create_pivot_and_inject(wb_to_fill, '收入汇总', full_df[full_df['科目'].isin(income_subjects)])
        _create_pivot_and_inject(wb_to_fill, '支出汇总', full_df[full_df['科目'].isin(expense_subjects)])

    except KeyError:
        logging.error("配置或预制件不完整，无法生成收支汇总表。")

def _create_pivot_and_inject(wb, sheet_name, df_filtered):
    if df_filtered.empty or sheet_name not in wb.sheetnames: return
    
    ws = wb[sheet_name]
    # 清空旧内容
    for row in ws.iter_rows():
        for cell in row: cell.value = None
            
    pivot = pd.pivot_table(df_filtered, values='金额', index='科目', columns='年份', aggfunc='sum').fillna(0)
    pivot['合计'] = pivot.sum(axis=1)
    pivot.loc['合计'] = pivot.sum()
    pivot = pivot.reset_index()

    # 注入新内容
    for r_idx, row in enumerate(pivot.itertuples(index=False), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)