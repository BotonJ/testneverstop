# /src/legacy_runner.py

import pandas as pd
import logging
from openpyxl import Workbook

logger = logging.getLogger(__name__)

def run_legacy_extraction(source_path: str, mapping_configs: dict):
    # ... (旧的提取逻辑，最终生成 raw_df) ...
    # 此处省略具体提取代码，假设它能成功生成raw_df
    from modules.balance_sheet_processor import process_balance_sheet
    from modules.income_statement_processor import process_income_statement
    from openpyxl import load_workbook as lw
    
    workbook = lw(source_path, data_only=True)
    all_data = []
    # ... (遍历sheet，调用处理器，追加到all_data的逻辑)
    alias_map_df = mapping_configs.get('科目等价映射')
    bs_map_df = mapping_configs.get('资产负债表区块')
    is_map_df = mapping_configs.get('业务活动表逐行')
    is_summary_config_df = mapping_configs.get('业务活动表汇总注入配置')

    for sheet_name in workbook.sheetnames:
        if '资产' in sheet_name or 'zcfz' in sheet_name.lower():
            # ... process_balance_sheet ...
            records = process_balance_sheet(workbook[sheet_name], sheet_name, bs_map_df, alias_map_df)
            if records: all_data.extend(records)
        elif '业务' in sheet_name or 'yewu' in sheet_name.lower():
            # ... process_income_statement ...
            records = process_income_statement(workbook[sheet_name], sheet_name, alias_map_df, is_map_df.to_dict('records'), is_summary_config_df)
            if records: all_data.extend(records)
    
    raw_df = pd.DataFrame(all_data)
    # --- 新增：创建内存中的预制件 ---
    prebuilt_wb = create_prebuilt_workbook(raw_df, mapping_configs)
    
    return raw_df, prebuilt_wb

def create_prebuilt_workbook(raw_df: pd.DataFrame, mapping_configs: dict):
    """
    【V2 - Bug修复】
    根据raw_df在内存中创建一个符合旧版格式的Excel预制件。
    修复了因硬编码导致业务活动表预制件创建失败的bug。
    """
    logger.info("  -> 正在内存中创建数据预制件...")
    wb = Workbook()
    wb.remove(wb.active) # 移除默认Sheet

    # 提取业务活动表配置，用于精确定位
    yewu_line_map = mapping_configs.get("业务活动表逐行")
    if yewu_line_map is None or yewu_line_map.empty:
        logger.warning("未能从配置中加载 '业务活动表逐行'，业务活动表预制件可能不完整。")
        yewu_line_map = pd.DataFrame()

    # 按年份分组
    for year, year_df in raw_df.groupby('年份'):
        # --- 创建资产负债表预制Sheet ---
        bs_df = year_df[year_df['报表类型'] == '资产负债表']
        if not bs_df.empty:
            ws_bs = wb.create_sheet(title=f"{year}资产负债表")
            # 填充 A/C/D 列 (左侧) 和 E/G/H 列 (右侧)
            # 为了简化，我们将所有项都先放在左侧，后续格式化函数会处理
            row_cursor = 2
            ws_bs.cell(row=1, column=1, value="项目")
            ws_bs.cell(row=1, column=3, value="期初金额")
            ws_bs.cell(row=1, column=4, value="期末金额")
            for _, row in bs_df.iterrows():
                ws_bs.cell(row=row_cursor, column=1, value=row['项目'])
                ws_bs.cell(row=row_cursor, column=3, value=row['期初金额'])
                ws_bs.cell(row=row_cursor, column=4, value=row['期末金额'])
                row_cursor += 1

        # --- 创建业务活动表预制Sheet ---
        is_df = year_df[year_df['报表类型'] == '业务活动表']
        if not is_df.empty:
            ws_is = wb.create_sheet(title=f"{year}业务活动表")
            
            # 使用 `yewu_line_map` 配置来精确填充预制件
            for _, config_row in yewu_line_map.iterrows():
                field_name = config_row.get("字段名")
                src_coord = config_row.get("源期末坐标") # 这是我们要填充到预制件的位置
                
                if not field_name or not src_coord:
                    continue
                    
                # 从当年的业务活动表数据(is_df)中安全地查找值
                value_series = is_df.loc[is_df['项目'] == field_name, '期末金额']
                
                # 如果找到了值，就填充到预制件的指定坐标
                if not value_series.empty:
                    value_to_fill = value_series.iloc[0]
                    try:
                        ws_is[src_coord] = value_to_fill
                    except Exception as e:
                        logger.warning(f"填充预制件单元格 {src_coord} 失败: {e}")
                # else: # 找不到就不填充，单元格将为空

    return wb