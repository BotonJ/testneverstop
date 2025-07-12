# /modules/mapping_loader.py
import pandas as pd
import openpyxl
from src.utils.logger_config import logger
from modules.utils import normalize_name

def load_mapping_file(path):
    """
    【最终版 V3 - 忠于原始逻辑】
    精确解析mapping_file，并增加对业务活动表汇总配置的解析。
    """
    logger.info("--- 开始使用原始逻辑精确解析 mapping_file.xlsx ---")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        logger.error(f"映射文件未找到: {path}")
        return {}

    all_mappings = {}

    def _load_sheet_as_df(sheet_name):
        if sheet_name in wb.sheetnames:
            return pd.read_excel(path, sheet_name=sheet_name)
        logger.warning(f"在mapping_file中未找到名为'{sheet_name}'的Sheet。")
        return pd.DataFrame()

    all_mappings["blocks_df"] = _load_sheet_as_df("资产负债表区块")
    all_mappings["alias_map_df"] = _load_sheet_as_df("科目等价映射")
    
    df_yewu = _load_sheet_as_df("业务活动表逐行")
    all_mappings["yewu_line_map"] = df_yewu.dropna(how='all').to_dict('records')

    yewu_subtotal_config = {}
    df_yewu_summary = _load_sheet_as_df("业务活动表汇总注入配置")
    if not df_yewu_summary.empty:
        if all(col in df_yewu_summary.columns for col in ['类型', '科目名称']):
            # 对科目名称进行清洗
            df_yewu_summary['科目名称'] = df_yewu_summary['科目名称'].apply(lambda x: normalize_name(x) if pd.notna(x) else "")
            grouped = df_yewu_summary.groupby('类型')['科目名称'].apply(list)
            yewu_subtotal_config = grouped.to_dict()
            logger.info(f"成功解析'业务活动表汇总注入配置'，识别出类型: {list(yewu_subtotal_config.keys())}")
        else:
            logger.error("'业务活动表汇总注入配置'缺少'类型'或'科目名称'列，无法用于复核。")
    all_mappings["yewu_subtotal_config"] = yewu_subtotal_config
    
    logger.info("--- mapping_file.xlsx 解析完成 ---")
    
    return all_mappings