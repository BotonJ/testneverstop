# /src/report_formatters/format_change_tables.py

import logging
import pandas as pd

logger = logging.getLogger(__name__)

def _parse_config_and_data(df_full: pd.DataFrame):
    """
    【新增辅助函数】
    从一个完整的配置Sheet DataFrame中，解析出头部的键值对配置和主体的数据映射部分。
    模拟旧版 get_mapping_conf_and_df 的功能。
    """
    if df_full is None or df_full.empty:
        return {}, pd.DataFrame()

    conf = {}
    data_start_row = 0
    
    # 寻找数据部分的起始行 (表头行)
    header_keywords = ["来源字段", "目标单元格（期初）", "区块名称", "起始行"]
    for i, row in df_full.iterrows():
        # 将行转换为字符串列表，以便于查找关键字
        row_values = [str(cell).strip() for cell in row if pd.notna(cell)]
        if any(keyword in row_values for keyword in header_keywords):
            data_start_row = i
            break
    
    # 解析头部的键值对配置
    # iloc[:data_start_row] 会选取数据起始行之前的所有行
    conf_df = df_full.iloc[:data_start_row]
    for _, row in conf_df.iterrows():
        key = row.iloc[0]
        value = row.iloc[1]
        if pd.notna(key) and pd.notna(value):
            conf[str(key).strip()] = str(value).strip()
            
    # 提取主体的数据映射DataFrame
    # iloc[data_start_row+1:] 会选取表头行之后的所有行
    # 并使用表头行的值作为新的列名
    data_df = df_full.iloc[data_start_row+1:].copy()
    data_df.columns = df_full.iloc[data_start_row].values
    data_df.reset_index(drop=True, inplace=True)
    
    return conf, data_df

def populate_balance_change_sheet(prebuilt_wb, wb_to_fill, mapping_configs):
    """
    【V2 - Bug修复】
    填充'资产负债变动'Sheet页的三个核心表格。
    修复了因错误解包DataFrame导致的ValueError。
    """
    logger.info("  -> 填充 '资产负债变动' Sheet...")
    try:
        ws_tgt = wb_to_fill['资产负债变动']
        
        # --- [修复] ---
        # 1. 先安全地获取完整的配置DataFrame
        df_inj1 = mapping_configs.get("inj1")
        df_inj2 = mapping_configs.get("inj2")
        df_inj3 = mapping_configs.get("inj3")
        
        # 2. 调用辅助函数来解析出 conf 和 df_map
        conf1, df1 = _parse_config_and_data(df_inj1)
        conf2, df2 = _parse_config_and_data(df_inj2)
        conf3, df3 = _parse_config_and_data(df_inj3)
        # --- [修复结束] ---
        
        # 按顺序注入
        _inject_table1(prebuilt_wb, ws_tgt, conf1, df1)
        _inject_table3(prebuilt_wb, ws_tgt, conf3, df3, mapping_configs) # Table2的逻辑暂时省略
        
    except KeyError:
        logger.error("模板中未找到'资产负债变动'Sheet或相关配置缺失。")
        
def _inject_table1(wb_src, ws_tgt, conf, df_map):
    """填充资产/负债/净资产总额变动表。"""
    if not conf or df_map is None or df_map.empty: return
    
    start_sheet_name = conf.get("start_sheet")
    end_sheet_name = conf.get("end_sheet")
    if not start_sheet_name or not end_sheet_name or not (start_sheet_name in wb_src.sheetnames and end_sheet_name in wb_src.sheetnames):
        logger.warning("Table1配置不完整或在预制件中找不到对应的Sheet，跳过注入。")
        return

    ws_start = wb_src[start_sheet_name]
    ws_end = wb_src[end_sheet_name]

    for _, row in df_map.iterrows():
        src_field = str(row.get("来源字段", "")).strip()
        if not src_field: continue
        
        # 使用更可靠的DataFrame进行查找，而不是遍历单元格
        start_row_data = ws_start[ws_start.iloc[:, 0] == src_field]
        end_row_data = ws_end[ws_end.iloc[:, 0] == src_field]

        val_init = start_row_data.iloc[0, 2] if not start_row_data.empty else 0 # 第3列是期初
        val_final = end_row_data.iloc[0, 3] if not end_row_data.empty else 0 # 第4列是期末
        
        tgt_init_cell = row.get("目标单元格（期初）")
        tgt_final_cell = row.get("目标单元格（期末）")
        tgt_var_cell = row.get("变动单元格")

        if pd.notna(tgt_init_cell): ws_tgt[tgt_init_cell] = val_init
        if pd.notna(tgt_final_cell): ws_tgt[tgt_final_cell] = val_final
        if pd.notna(tgt_var_cell): ws_tgt[tgt_var_cell] = val_final - val_init

def _inject_table3(wb_src, ws_tgt, conf, df_map, mapping_configs):
    """填充净资产构成表并应用公式。"""
    if not conf or df_map is None or df_map.empty: return

    start_sheet_name = conf.get("start_sheet")
    end_sheet_name = conf.get("end_sheet")
    if not start_sheet_name or not end_sheet_name or not (start_sheet_name in wb_src.sheetnames and end_sheet_name in wb_src.sheetnames):
        logger.warning("Table3配置不完整或在预制件中找不到对应的Sheet，跳过注入。")
        return
    
    ws_start = wb_src[start_sheet_name]
    ws_end = wb_src[end_sheet_name]

    for _, row in df_map.iterrows():
        if pd.isna(row.get("来源字段")): continue
        
        val_start = ws_start[row["来源单元格（期初）"]].value or 0
        val_end = ws_end[row["来源单元格（期末）"]].value or 0
        change = val_end - val_start
        
        if pd.notna(row.get("目标单元格（期初）")): ws_tgt[row["目标单元格（期初）"]] = val_start
        if pd.notna(row.get("目标单元格（期末）")): ws_tgt[row["目标单元格（期末）"]] = val_end

        if change > 0 and pd.notna(row.get("增加单元格")):
            ws_tgt[row["增加单元格"]] = change
        elif change < 0 and pd.notna(row.get("减少单元格")):
            ws_tgt[row["减少单元格"]] = abs(change)
            
    # 应用公式
    try:
        df_formula = mapping_configs.get("合计公式配置")
        if df_formula is not None:
            for _, row in df_formula.iterrows():
                if pd.notna(row.get("变动单元格")) and pd.notna(row.get("变动公式")):
                    ws_tgt[row["变动单元格"]] = f'={row["变动公式"]}'
    except Exception:
        pass