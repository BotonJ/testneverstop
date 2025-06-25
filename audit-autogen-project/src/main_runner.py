import os
import logging
from openpyxl import load_workbook
from modules.mapping_loader import load_mapping_file
from modules.fill_yewu import fill_yewu_by_mapping
#from modules.fill_utils import fill_balance_block
from modules.fill_balance_anchor import fill_balance_sheet_by_name
from modules.render_header import render_header
from config import DATA_DIR, OUTPUT_DIR, LOG_DIR  # ✅ 引入路径配置
from modules.log_utils import log_write

os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(filename=os.path.join(LOG_DIR, "auto_log.txt"),
                    level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
log_balance = []# 初始化日志（可以放循环外或内） 
log_yewu = []

def get_balance_net_assets(ws_balance):
    for i in range(1, ws_balance.max_row + 1):
        name = ws_balance[f"A{i}"].value
        if name and "净资产合计" in str(name):
            val_initial = ws_balance[f"B{i}"].value or 0
            val_final = ws_balance[f"C{i}"].value or 0
            return {"期初": val_initial, "期末": val_final}
    return {"期初": 0, "期末": 0}

def main():
    try:
        # ✅ 路径拼接
        mapping_path = os.path.join(DATA_DIR, "mapping_file.xlsx")
        wb_src_path = os.path.join(DATA_DIR, "soce.xlsx")
        wb_tgt_path = os.path.join(DATA_DIR, "t.xlsx")
        output_path = os.path.join(OUTPUT_DIR, "output.xlsx")
        
        # ✅ 加载数据
        mapping = load_mapping_file(mapping_path)
        alias_dict = {a: k for k, v in mapping["subject_alias_map"].items() for a in [k] + v}
        wb_src = load_workbook(wb_src_path, data_only=True)
        wb_tgt = load_workbook(wb_tgt_path)

        prev_ws_yewu = None

        for sheet_name in wb_src.sheetnames:
            if "资产负债表" in sheet_name:
                year = int(sheet_name[:4])
                logging.info(f"▶️ 正在处理：{sheet_name}")
                ws_src = wb_src[sheet_name]

                ws_balance = wb_tgt.copy_worksheet(wb_tgt["资产负债表"])
                ws_balance.title = f"{year}资产负债表"            
                fill_balance_sheet_by_name(ws_src, ws_balance, alias_dict, log_balance, skip_list=[])
                logging.info(f"✅ 已注入：{year}资产负债表 → {ws_balance.title}")
                render_header(wb_tgt, sheet_name=f"{year}资产负债表", year=year, header_meta=mapping["header_meta"])

                if f"{year}业务活动表" in wb_src.sheetnames:
                    ws_src_yewu = wb_src[f"{year}业务活动表"]
                    ws_yewu = wb_tgt.copy_worksheet(wb_tgt["业务活动表"])
                    ws_yewu.title = f"{year}业务活动表"
                    net_asset_fallback = get_balance_net_assets(ws_balance)  # 提取资产负债表净资产合计
                    fill_yewu_by_mapping(
                        ws_src_yewu, ws_yewu, mapping["yewu_line_map"], 
                        prev_ws=prev_ws_yewu,net_asset_fallback=net_asset_fallback,log=log_yewu
                    )
                    logging.info(f"✅ 已注入：{year}业务活动表 → {ws_yewu.title}")
                    render_header(wb_tgt, sheet_name=f"{year}业务活动表", year=year, header_meta=mapping["header_meta"])
                    prev_ws_yewu = ws_yewu
        
        for tmpl_sheet in ["资产负债表", "业务活动表"]:
            if tmpl_sheet in wb_tgt.sheetnames:
                wb_tgt.remove(wb_tgt[tmpl_sheet])        
        with open(os.path.join(LOG_DIR, "balance_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(log_balance))
        with open(os.path.join(LOG_DIR, "yewu_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(log_yewu))
        wb_tgt.save(output_path)        
        logging.info(f"✅ 报表生成完毕：{output_path}")
        logging.info("✅ 全部处理完成")
    except Exception as e:
        import traceback
        logging.error(f"❌ 程序运行异常：{e}")
        logging.error(traceback.format_exc())
    
if __name__ == "__main__":
    main()


