# /modules/income_statement_processor.py
import re
from src.utils.logger_config import logger

def process_income_statement(ws_src, sheet_name, yewu_line_map, alias_map_df):
    """
    【最终版 - 忠于原始逻辑】
    模拟 fill_yewu.py 的“精确坐标，点对点复制”算法。
    """
    logger.info(f"--- 开始处理业务活动表: '{sheet_name}' (使用原始'精确坐标'逻辑) ---")
    if not yewu_line_map:
        logger.warning(f"跳过Sheet '{sheet_name}'，因为'yewu_line_map'配置为空。")
        return []

    # --- 准备工作: 我们依然需要合计项的别名来做判断 ---
    income_total_aliases = ['收入合计', '一、收 入', '（一）收入合计']
    expense_total_aliases = ['费用合计', '二、费 用', '（二）费用合计']

    records = []
    year = (re.search(r'(\d{4})', sheet_name) or [None, "未知"])[1]

    # --- 主流程: 遍历`业务活动表逐行`配置，进行点对点提取 ---
    for item in yewu_line_map:
        subject_name = item.get("字段名")
        src_initial_coord = item.get("源期初坐标")
        src_final_coord = item.get("源期末坐标")

        if not subject_name or not src_initial_coord or not src_final_coord:
            continue
        
        subject_name = subject_name.strip()

        # --- 提取数据 ---
        try:
            # 使用openpyxl的安全方式读取单元格
            start_val = ws_src[src_initial_coord].value
            end_val = ws_src[src_final_coord].value
        except Exception as e:
            logger.error(f"在提取 '{subject_name}' 时读取坐标失败: {e}。坐标: 初'{src_initial_coord}', 末'{src_final_coord}'")
            continue

        # --- 判断科目类型 ---
        standard_name = subject_name
        subject_type = '普通'
        if subject_name in income_total_aliases:
            standard_name = '收入合计'
            subject_type = '合计'
        elif subject_name in expense_total_aliases:
            standard_name = '费用合计'
            subject_type = '合计'

        record = {
            "来源Sheet": sheet_name,
            "报表类型": "业务活动表",
            "年份": year,
            "项目": standard_name,
            "科目类型": subject_type,
            "本期金额": end_val,   # 源期末坐标对应本期
            "上期金额": start_val, # 源期初坐标对应上期
        }
        records.append(record)

    logger.info(f"--- 业务活动表 '{sheet_name}' 处理完成，生成 {len(records)} 条记录。---")
    return records