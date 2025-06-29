# main_runner.py：主流程整合脚本
import os
from pathlib import Path
from openpyxl import load_workbook
from modules.collector import collect_summary_values
from src.legacy_runner import run_main_injection 
from inject_modules.table_injector import inject_tables_and_summary
from inject_modules.text_renderer import render_text_template_from_mapping, inject_text_to_excel, inject_summary_values_debug
from modules.mapping_loader import load_mapping_file
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def run_main():
    project_root = Path(__file__).resolve().parents[1]
    mapping_path = project_root / "data" / "mapping_file.xlsx"
    source_path = project_root / "output" / "output.xlsx"
    template_path = project_root / "data" / "template_file.xlsx"
    final_path = project_root / "output" / "final_report.xlsx"

    os.makedirs(project_root / "output", exist_ok=True)
    logging.info("Starting report generation process...")

    # ✅ 构造 alias_dict
    raw_alias_map = load_mapping_file(mapping_path)["subject_alias_map"]
    alias_dict = {}
    for std, aliases in raw_alias_map.items():
        std = std.strip()
        # ✅ 统一转为列表（无论 aliases 是 str 还是 list）
        if not isinstance(aliases, list):
            aliases = [aliases]
        for alias in [std] + aliases:
            alias_dict[alias.strip()] = std

    # Step 0：如果 output.xlsx 不存在，则运行 legacy_runner
    if not source_path.exists():
        logging.info("Output.xlsx 未找到，启动 legacy_runner 自动注入流程...")
        print("Output.xlsx 未找到，启动 legacy_runner 自动注入流程...")
        run_main_injection()
        if not source_path.exists():
            logging.error("运行 legacy_runner 后仍未找到 output.xlsx，终止执行。")
            print("运行 legacy_runner 后仍未找到 output.xlsx，终止执行。")
            return      

    # ✅ 第一步：提取 summary_values
    logging.info("Step 1: Collecting summary values...")
    summary_values = collect_summary_values(mapping_path, source_path)
    logging.info(f"Collected Summary Values: {summary_values}")

    # ✅ 补全 summary_values 中的标准字段
    if alias_dict:
        extended = {}
        for k, v in summary_values.items():
            std_key = alias_dict.get(k)
            if std_key and std_key not in summary_values:
                extended[std_key] = v
        summary_values.update(extended)

    # ✅ 第二步：插入三张表和字段 → 保存 final 报表
    logging.info("Step 2: Injecting tables and summary into final report...")
    try:
        inject_tables_and_summary(
            output_path=source_path,
            template_path=template_path,
            summary_values=summary_values,
            dest_path=final_path,
            alias_dict=alias_dict
        )
        logging.info("Tables and summary injected successfully.")
    except Exception as e:
        logging.error(f"Error during table and summary injection: {e}")
        return

    # ✅ 第三步：渲染文字模板并写入 final
    logging.info("Step 3: Rendering text template and injecting into final report...")
    rendered_text = render_text_template_from_mapping(mapping_path, summary_values, alias_dict)
    logging.info(f"Rendered Text (first 200 chars): {rendered_text[:200]}...")
    try:
        inject_text_to_excel(final_path, sheet_name="汇总区块", cell="K1", text=rendered_text)
        logging.info("Text template rendered and injected successfully.")
    except Exception as e:
        logging.error(f"Error during text injection: {e}")

    # ✅ 第四步：调试用途写入 summary_values 到元数据 Sheet
    logging.info("Step 4: Injecting summary values to debug sheet...")
    try:
        inject_summary_values_debug(final_path, summary_values)
        logging.info("Debug summary values injected successfully.")
    except Exception as e:
        logging.error(f"Error during debug summary values injection: {e}")

    logging.info(f"✅ 报表已完成写入：{final_path}")
    logging.info("Report generation process completed.")

if __name__ == '__main__':
    run_main()
