# /main.py
import sys
import os
import json
from src.utils.logger_config import logger
from src.legacy_runner import run_legacy_extraction
from src.data_processor import pivot_and_clean_data, calculate_summary_values
from src.data_validator import run_all_checks
from src.report_writer import ReportWriter # <-- 导入新的报告生成器
from modules.mapping_loader import load_mapping_file

SRC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
if SRC_PATH not in sys.path:
    sys.path.append(SRC_PATH)

def run_audit_report():
    logger.info("========================================")
    logger.info("===    自动化审计报告生成流程启动    ===")
    logger.info("========================================")

    project_root = os.path.dirname(os.path.abspath(__file__))
    source_file = os.path.join(project_root, 'data', 'soce.xlsx')
    mapping_file = os.path.join(project_root, 'data', 'mapping_file.xlsx')
    output_file = os.path.join(project_root, 'output', 'final_audit_report.xlsx') # <-- 定义最终输出文件名
    os.makedirs(os.path.join(project_root, 'output'), exist_ok=True) # 确保输出目录存在
    
    logger.info(f"源文件路径: {source_file}")
    logger.info(f"映射文件路径: {mapping_file}")
    logger.info(f"最终报告输出路径: {output_file}")

    # --- 步骤 1/5: 数据提取 ---
    logger.info("\n--- [步骤 1/5] 执行数据提取 ---")
    raw_df = run_legacy_extraction(source_file, mapping_file)
    if raw_df is None or raw_df.empty: return
    logger.info("✅ 数据提取成功！")

    # --- 步骤 2/5: 数据处理与计算 ---
    logger.info("\n--- [步骤 2/5] 执行数据处理与计算 ---")
    pivoted_normal_df, pivoted_total_df = pivot_and_clean_data(raw_df)
    if pivoted_total_df is None or pivoted_total_df.empty: return
    logger.info("✅ 数据透视与清理成功！")
        
    final_summary_dict = calculate_summary_values(pivoted_total_df, raw_df)
    if not final_summary_dict: return
    logger.info("✅ 最终汇总指标计算成功！")
    
    # --- 步骤 3/5: 执行数据复核 ---
    logger.info("\n--- [步骤 3/5] 执行数据复核 ---")
    full_mapping = load_mapping_file(mapping_file)
    verification_results = run_all_checks(pivoted_normal_df, pivoted_total_df, raw_df, full_mapping)
    logger.info("✅ 数据复核完成！")

    # --- 步骤 4/5: 生成最终Excel报告 ---
    logger.info("\n--- [步骤 4/5] 执行报告生成与注入 ---")
    writer = ReportWriter(output_file)
    writer.generate_full_report(
        summary_dict=final_summary_dict,
        verification_results=verification_results,
        normal_df=pivoted_normal_df,
        total_df=pivoted_total_df,
        mapping=full_mapping
    )
    logger.info("✅ 报告生成完毕！")

    # --- 步骤 5/5: 在终端展示简报 ---
    logger.info("\n--- [步骤 5/5] 在终端展示简报 ---")
    print("\n" + "="*25 + " 最终计算结果 " + "="*25)
    print(json.dumps(final_summary_dict, indent=4, ensure_ascii=False))
    print("="*68)
    print("\n" + "="*27 + " 复核报告摘要 " + "="*27)
    for line in verification_results: print(line)
    print("="*68)
    
    logger.info("\n========================================")
    logger.info("===         流程执行完毕           ===")
    logger.info("========================================")

if __name__ == '__main__':
    run_audit_report()