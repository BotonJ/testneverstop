# collector.py：统一提取 summary_values
import json
import logging # Added for logging
from openpyxl import load_workbook
from modules.mapping_loader import load_mapping_file
from inject_modules.balance_utils import get_balance_core_data

# Configure logging for debugging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def collect_summary_values(mapping_path, output_path):
    """
    从 mapping_file 的 HeaderMapping 中提取单位名、起止sheet；
    从 output.xlsx 中读取资产/负债/净资产期初期末值，并计算差额；
    返回完整 summary_values 字典
    """
    summary = {}
    mapping = load_mapping_file(mapping_path)
    try:
        mapping_wb = load_workbook(mapping_path, data_only=True)
        header_ws = mapping_wb["HeaderMapping"]

        # 获取 HeaderMapping 规则字段
        rule_dict = {
            row[0].value: str(row[2].value).strip() if row[2].value is not None else ""
            for row in header_ws.iter_rows(min_row=2)
        }

        summary["单位名称"] = rule_dict.get("单位名称", "【未提取】")
        summary["审计期间"] = rule_dict.get("期末", "【未提取】")

        start_sheet = rule_dict.get("起始资产负债表Sheet")
        end_sheet = rule_dict.get("终止资产负债表Sheet")

        # Log extracted sheet names for debugging
        logging.info(f"HeaderMapping - 起始资产负债表Sheet: {start_sheet}")
        logging.info(f"HeaderMapping - 终止资产负债表Sheet: {end_sheet}")

        # 从 output.xlsx 提取资产负债字段
        wb = load_workbook(output_path, data_only=True)
        if start_sheet in wb.sheetnames and end_sheet in wb.sheetnames:
            block_map = mapping["blocks"]
            print("🧾 当前 block_map 内容如下：")
            print(json.dumps(block_map, ensure_ascii=False, indent=2))                
            start_data = get_balance_core_data(wb[start_sheet], block_map, mapping["subject_alias_map"])
            end_data = get_balance_core_data(wb[end_sheet], block_map, mapping["subject_alias_map"])
            

            # Log core data for debugging
            logging.info(f"Start Sheet Data ({start_sheet}): {start_data}")
            logging.info(f"End Sheet Data ({end_sheet}): {end_data}")

            for field in ["资产总额", "负债总额", "净资产总额"]:
                start_val = float(start_data.get(f"期初{field}", 0) or 0)
                end_val = float(end_data.get(f"期末{field}", 0) or 0)

                summary[f"期初{field}"] = start_val
                summary[f"期末{field}"] = end_val
                try:
                    summary[f"{field}增减"] = round(end_val - start_val, 2)
                except Exception as e:
                    summary[f"{field}增减"] = f"计算失败: {e}"
                try:
                    diff = round(end_val - start_val, 2)
                    summary[f"{field}增减"] = diff
                except Exception as e:
                    summary[f"{field}增减"] = f"计算失败: {e}"
                    logging.error(f"Error calculating {field} difference: {e}")
        else:
            summary["期初资产总额"] = summary["期末资产总额"] = "【未找到起止Sheet】"
            summary["期初负债总额"] = summary["期末负债总额"] = "【未找到起止Sheet】"
            summary["期初净资产总额"] = summary["期末净资产总额"] = "【未找到起止Sheet】"
            summary["资产总额增减"] = summary["负债总额增减"] = summary["净资产总额增减"] = "【未找到起止Sheet】"
            logging.warning(f"Required sheets not found in {output_path}. Start: '{start_sheet}', End: '{end_sheet}'.")

    except Exception as e:
        logging.error(f"Error in collect_summary_values: {e}")
        # Populate summary with error indicators in case of a global failure
        for key in ["单位名称", "审计期间", "期初资产总额", "期末资产总额", "资产总额增减",
                     "期初负债总额", "期末负债总额", "负债总额增减",
                     "期初净资产总额", "期末净资产总额", "净资产总额增减"]:
            if key not in summary:
                summary[key] = f"【提取失败: {e}】"
    return summary
    
    