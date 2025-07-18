# /src/report_generator.py

import logging
import openpyxl
from openpyxl.styles import Alignment
from openpyxl import Workbook
import pandas as pd
import re
import shutil

# 从新的格式化模块导入函数
from src.report_formatters.format_balance import format_balance_sheet
from src.report_formatters.format_yewu import format_yewu_sheet
from src.report_formatters.format_change_tables import populate_balance_change_sheet
from src.report_formatters.format_biz_summary import create_and_inject_biz_summary

logger = logging.getLogger(__name__)


def _render_header_on_sheet(ws_tgt, year, audit_end_year, header_config):
    """
    【V3 - 最终版渲染逻辑】
    根据用户定义的复杂规则，为一个worksheet对象渲染动态表头。
    """
    if not header_config: return
    try:
        for field_name, meta in header_config.items():
            target_cell, value_to_write = meta.get("target_cell"), None
            if not target_cell: continue

            # 根据您提供的逻辑表格进行渲染
            if "资产负债表" in ws_tgt.title:
                if field_name == "期初": value_to_write = f"{year - 1}年12月31日"
                elif field_name == "期末": value_to_write = f"{year}年12月31日"
                
            elif "业务活动表" in ws_tgt.title:
                if field_name == "期初": value_to_write = f"{year - 1}年累计数"
                elif field_name == "期末":
                    if year == audit_end_year: value_to_write = f"{year}年1-3月累计数"
                    else: value_to_write = f"{year}年累计数"
            
            # 处理其他如“单位名称”等通用字段
            if field_name == "单位名称":
                value_to_write = meta.get("rule")

            if value_to_write:
                # 处理多单元格写入 (如 A3,A35)
                for cell in str(target_cell).split(','):
                    ws_tgt[cell.strip()] = value_to_write

    except Exception as e:
        logger.warning(f"为 Sheet '{ws_tgt.title}' 渲染表头时失败: {e}")


def apply_global_formatting(wb):
    """应用专业的全局数字格式。"""
    logger.info("  -> 应用全局数字格式...")
    summary_format, activity_format = '#,##0.00;-#,##0.00;"-"', '#,##0.00;-#,##0.00;;@'
    right_align = Alignment(horizontal='right', vertical='center')
    sheet_names = ['资产负债变动', '收入汇总', '支出汇总'] + [s.title for s in wb if re.match(r"^\d{4}_(资产负债表|业务活动表)$", s.title)]
    for name in sheet_names:
        if name not in wb.sheetnames: continue
        ws, formatter = wb[name], activity_format if "业务活动表" in name else summary_format
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format, cell.alignment = formatter, right_align


def _create_prebuilt_workbook_in_memory(raw_df: pd.DataFrame, mapping_configs: dict):
    """【内部函数】根据raw_df在内存中创建数据预制件。"""
    logger.info("  -> 正在内存中创建数据预制件...")
    wb = Workbook()
    wb.remove(wb.active)
    yewu_map = mapping_configs.get("业务活动表逐行", pd.DataFrame())

    for year, year_df in raw_df.groupby('年份'):
        # 创建资产负债表预制Sheet
        bs_df = year_df[year_df['报表类型'] == '资产负债表']
        if not bs_df.empty:
            ws_bs = wb.create_sheet(title=f"{year}资产负债表")
            ws_bs.append(["项目", "期初金额", "期末金额"]) # 只是为了方便调试，实际不使用
            for _, row in bs_df.iterrows():
                # 为了兼容旧版table_injector, 我们需要模拟A/C/D列格式
                ws_bs.append([row['项目'], row['期初金额'], row['期末金额']])

        # 创建业务活动表预制Sheet
        is_df = year_df[year_df['报表类型'] == '业务活动表']
        if not is_df.empty and not yewu_map.empty:
            ws_is = wb.create_sheet(title=f"{year}业务活动表")
            for _, cfg_row in yewu_map.iterrows():
                field, coord = cfg_row.get("字段名"), cfg_row.get("源期末坐标")
                if not field or not coord: continue
                val_series = is_df.loc[is_df['项目'] == field, '期末金额']
                if not val_series.empty: ws_is[coord] = val_series.iloc[0]
    return wb

def generate_master_report(raw_df, summary_values, mapping_configs, template_path, output_path):
    """
    【V5.4 终极版】
    - 修复重复Sheet的BUG。
    - 实现最终版的复杂表头渲染逻辑。
    """
    logger.info("--- [报告生成模块启动] ---")

    # 1. 准备报告文件：在模块内部复制模板
    try:
        shutil.copy(template_path, output_path)
    except Exception as e:
        logger.error(f"复制模板文件时出错: {e}"); return

    # 2. 在内存中创建数据预制件
    prebuilt_wb = _create_prebuilt_workbook_in_memory(raw_df, mapping_configs)

    # 3. 加载刚刚复制好的报告文件，准备写入
    wb_final = openpyxl.load_workbook(output_path)
        
    # 4. 创建年度报表
    logger.info("-> 步骤 A: 创建并格式化年度报表 Sheets...")
    alias_dict = mapping_configs.get("alias_dict", {})
    yewu_map_df = mapping_configs.get("业务活动表逐行", pd.DataFrame())
    
    header_meta = mapping_configs.get("HeaderMapping", {})
    header_balance_config = header_meta.get("资产负债表", {})
    header_yewu_config = header_meta.get("业务活动表", {})
    
    years = sorted(raw_df['年份'].unique())
    start_year, end_year = years[0], years[-1]
    prev_yewu_ws = None

    for year in years:
        # --- [修复] 调整循环和调用逻辑 ---
        # --- 处理资产负债表 ---
        sheet_name_balance_src = f"{year}资产负债表"
        if sheet_name_balance_src in prebuilt_wb.sheetnames:
            ws_src_balance = prebuilt_wb[sheet_name_balance_src]
            ws_tgt_balance = wb_final.copy_worksheet(wb_final["资产负债表"])
            ws_tgt_balance.title = f"{year}_资产负债表"
            format_balance_sheet(ws_src_balance, ws_tgt_balance, alias_dict)
            _render_header_on_sheet(ws_tgt_balance, year, end_year, header_balance_config)
        
        # --- 处理业务活动表 ---
        sheet_name_yewu_src = f"{year}业务活动表"
        if sheet_name_yewu_src in prebuilt_wb.sheetnames:
            ws_src_yewu = prebuilt_wb[sheet_name_yewu_src]
            net_change = summary_values.get(f"{year}_净资产变动额")
            ws_tgt_yewu = wb_final.copy_worksheet(wb_final["业务活动表"])
            ws_tgt_yewu.title = f"{year}_业务活动表"
            format_yewu_sheet(ws_src_yewu, ws_tgt_yewu, yewu_map_df, prev_ws=prev_yewu_ws, net_asset_change=net_change)
            _render_header_on_sheet(ws_tgt_yewu, year, end_year, header_yewu_config)
            prev_yewu_ws = ws_tgt_yewu

    # 5. 填充分析报表
    logger.info("-> 步骤 B: 填充分析与汇总 Sheets...")
    populate_balance_change_sheet(prebuilt_wb, wb_final, mapping_configs)
    create_and_inject_biz_summary(prebuilt_wb, wb_final, mapping_configs)
    
    # 6. 最终修饰
    logger.info("-> 步骤 C: 应用全局格式化并清理...")
    apply_global_formatting(wb_final)
    if '资产负债表' in wb_final.sheetnames: wb_final.remove(wb_final['资产负债表'])
    if '业务活动表' in wb_final.sheetnames: wb_final.remove(wb_final['业务活动表'])
    
    # 7. 保存
    try:
        wb_final.save(output_path)
        logger.info(f"✅✅✅ 最终审计报告已成功生成: {output_path}")
    except Exception as e:
        logger.error(f"保存最终报告时失败: {e}", exc_info=True) 