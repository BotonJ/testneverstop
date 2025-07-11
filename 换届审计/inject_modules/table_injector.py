# inject_modules/table_injector.py

from openpyxl.workbook import Workbook
from inject_modules.table1 import inject_table1
from inject_modules.table2 import inject_table2
from inject_modules.table3 import inject_table3
from inject_modules.mapping import get_mapping_conf_and_df
import logging



def populate_balance_change_sheet(wb_src: Workbook, wb_to_fill: Workbook, mapping_file_path: str):
    """
    在目标工作簿(wb_to_fill)中，找到“资产负债变动”Sheet，并用源工作簿(wb_src)中的数据填充它。
    """
    target_sheet_name = "资产负债变动"
    if target_sheet_name not in wb_to_fill.sheetnames:
        logging.error(f"工作簿中未找到名为 '{target_sheet_name}' 的Sheet，无法注入数据。")
        return

    ws_tgt = wb_to_fill[target_sheet_name]
    logging.info(f"成功定位到目标Sheet: '{target_sheet_name}'")

    # 构造 df_map 配置
    conf1, df1 = get_mapping_conf_and_df(mapping_file_path, "inj1")
    conf2, df2 = get_mapping_conf_and_df(mapping_file_path, "inj2")
    conf3, df3 = get_mapping_conf_and_df(mapping_file_path, "inj3")

    # 注入三张表格到这个指定的Sheet
    logging.info(f"开始向 '{target_sheet_name}' Sheet注入Table 1...")
    inject_table1(wb_src, ws_tgt, conf1, df1, log=None)
    
    logging.info(f"开始向 '{target_sheet_name}' Sheet注入Table 2...")
    inject_table2(wb_src, ws_tgt, conf2, df2, log=None)
    
    logging.info(f"开始向 '{target_sheet_name}' Sheet注入Table 3...")
    inject_table3(wb_src, ws_tgt, conf3, df3, mapping_file_path, log=None)
    
    logging.info(f"'{target_sheet_name}' Sheet数据注入完成。")