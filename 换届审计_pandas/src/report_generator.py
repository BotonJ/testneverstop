# /src/report_generator.py

import logging
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
import re

# 从新的格式化模块导入函数
from .report_formatters.format_balance import format_balance_sheet
from .report_formatters.format_yewu import format_yewu_sheet
from .report_formatters.format_change_tables import populate_balance_change_sheet
from .report_formatters.format_biz_summary import create_and_inject_biz_summary

logger = logging.getLogger(__name__)


def _render_header_on_sheet(ws_tgt, year, header_config):
    """
    【V2 - 内部函数】为一个指定的worksheet对象渲染动态表头。
    :param ws_tgt: 目标worksheet对象
    :param year: 当前年份
    :param header_config: 针对该报表类型的表头配置
    """
    if not header_config: return
    try:
        for field_name, meta in header_config.items():
            value_to_write = None
            rule = meta.get("rule")
            target_cell = meta.get("target_cell")

            if not rule or not target_cell: continue

            if field_name == "单位名称":
                value_to_write = rule
            elif field_name == "报表日期":
                if "资产负债表" in ws_tgt.title:
                    value_to_write = rule.replace("%Y", str(year)).replace("%m", "12").replace("%d", "31")
                elif "业务活动表" in ws_tgt.title:
                    value_to_write = rule.replace("%Y", str(year))
            
            if value_to_write:
                ws_tgt[target_cell] = value_to_write
    except Exception as e:
        logger.warning(f"为 Sheet '{ws_tgt.title}' 渲染表头时失败: {e}")


def apply_global_formatting(wb):
    """【移植】应用专业的全局数字格式。"""
    logger.info("  -> 应用全局数字格式...")
    summary_format = '#,##0.00;-#,##0.00;"-"'
    activity_format = '#,##0.00;-#,##0.00;;@'
    right_alignment = Alignment(horizontal='right', vertical='center')
    
    # 获取所有需要格式化的sheet名称
    sheet_names_to_format = ['资产负债变动', '收入汇总', '支出汇总']
    for sheet in wb.worksheets:
        title = sheet.title
        # 匹配 "YYYY_资产负债表" 或 "YYYY_业务活动表"
        if re.match(r"^\d{4}_(资产负债表|业务活动表)$", title):
            sheet_names_to_format.append(title)
            
    for sheet_name in sheet_names_to_format:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        formatter = activity_format if "业务活动表" in sheet_name else summary_format
        for row in ws.iter_rows(min_row=2): # 从第二行开始，避免格式化标题
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = formatter
                    cell.alignment = right_alignment


def generate_master_report(prebuilt_wb, mapping_configs, output_path, summary_values):
    """
    【V5.1 精修版】
    修复重复Sheet的BUG，并正确注入动态表头。
    """
    logger.info("--- [最终步骤] 开始生成统一的主审计报告 ---")
    
    # 1. 复制模板作为最终报告的基础
    template_path = mapping_configs['template_path']
    try:
        wb_final = openpyxl.load_workbook(template_path)
    except FileNotFoundError:
        logger.error(f"模板文件未找到: {template_path}，无法生成报告。")
        return

    # 2. 创建年度报表 (调用格式化工具)
    logger.info("-> 步骤 A: 创建并格式化年度报表 Sheets...")
    alias_dict = mapping_configs.get("alias_dict", {})
    yewu_map_df = mapping_configs.get("业务活动表逐行")
    if yewu_map_df is None: yewu_map_df = pd.DataFrame() # 确保是DataFrame
    
    header_meta = mapping_configs.get("HeaderMapping", {})
    header_balance_config = header_meta.get("资产负债表", {})
    header_yewu_config = header_meta.get("业务活动表", {})
    
    years = sorted([int(s.title[:4]) for s in prebuilt_wb.worksheets if s.title[:4].isdigit()])
    prev_yewu_ws = None

    # --- [修复] 调整循环逻辑，确保不重复 ---
    for year in years:
        # --- 处理资产负债表 ---
        sheet_name_balance_src = f"{year}资产负债表"
        if sheet_name_balance_src in prebuilt_wb.sheetnames:
            ws_src_balance = prebuilt_wb[sheet_name_balance_src]
            
            # 复制模板并重命名
            ws_tgt_balance = wb_final.copy_worksheet(wb_final["资产负债表"])
            ws_tgt_balance.title = f"{year}_资产负债表"
            
            # 填充数据
            format_balance_sheet(ws_src_balance, ws_tgt_balance, alias_dict)
            
            # 注入表头
            _render_header_on_sheet(ws_tgt_balance, year, header_balance_config)
        
        # --- 处理业务活动表 ---
        sheet_name_yewu_src = f"{year}业务活动表"
        if sheet_name_yewu_src in prebuilt_wb.sheetnames:
            ws_src_yewu = prebuilt_wb[sheet_name_yewu_src]
            net_asset_change = summary_values.get(f"{year}_净资产变动额")
            
            # 复制模板并重命名
            ws_tgt_yewu = wb_final.copy_worksheet(wb_final["业务活动表"])
            ws_tgt_yewu.title = f"{year}_业务活动表"

            # 填充数据，并处理衔接
            format_yewu_sheet(ws_src_yewu, ws_tgt_yewu, yewu_map_df, prev_ws=prev_yewu_ws, net_asset_change=net_asset_change)
            
            # 注入表头
            _render_header_on_sheet(ws_tgt_yewu, year, header_yewu_config)
            
            # 更新 prev_ws 以供下一年使用
            prev_yewu_ws = ws_tgt_yewu

    # 3. 填充分析报表 (调用格式化工具)
    logger.info("-> 步骤 B: 填充分析与汇总 Sheets...")
    populate_balance_change_sheet(prebuilt_wb, wb_final, mapping_configs)
    create_and_inject_biz_summary(prebuilt_wb, wb_final, mapping_configs)
    
    # 4. 最终修饰
    logger.info("-> 步骤 C: 应用全局格式化并清理...")
    apply_global_formatting(wb_final)
    # 移除模板
    if '资产负债表' in wb_final.sheetnames: wb_final.remove(wb_final['资产负债表'])
    if '业务活动表' in wb_final.sheetnames: wb_final.remove(wb_final['业务活动表'])
    
    # 5. 保存
    try:
        wb_final.save(output_path)
        logger.info(f"✅✅✅ 最终审计报告已成功生成: {output_path}")
    except Exception as e:
        logger.error(f"保存最终报告时失败: {e}", exc_info=True)