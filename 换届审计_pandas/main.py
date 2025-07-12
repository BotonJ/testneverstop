# /main.py
import sys
import os
import json
from src.utils.logger_config import logger
from src.legacy_runner import run_legacy_extraction
from src.data_processor import pivot_and_clean_data, calculate_summary_values

SRC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
sys.path.append(SRC_PATH)

def run_audit_report():
    logger.info("========================================")
    logger.info("===    自动化审计报告生成流程启动    ===")
    logger.info("========================================")

    project_root = os.path.dirname(os.path.abspath(__file__))
    source_file = os.path.join(project_root, 'data', 'soce.xlsx')
    mapping_file = os.path.join(project_root, 'data', 'mapping_file.xlsx')
    
    logger.info(f"源文件路径: {source_file}")
    logger.info(f"映射文件路径: {mapping_file}")

    logger.info("\n--- [步骤 1/3] 执行数据提取 ---")
    raw_df = run_legacy_extraction(source_file, mapping_file)

    if raw_df is None or raw_df.empty:
        logger.error("数据提取失败或未提取到任何数据，流程终止。")
        return

    logger.info("✅ 数据提取成功！原始DataFrame已加载到内存。")

    logger.info("\n--- [步骤 2/3] 执行数据处理与计算 ---")
    
    pivoted_normal_df, pivoted_total_df = pivot_and_clean_data(raw_df)
    if pivoted_total_df is None or pivoted_total_df.empty:
        logger.error("数据透视后未能生成合计项目表，无法进行汇总计算，流程终止。")
        return
    logger.info("✅ 数据透视与清理成功！")
        
    final_summary_dict = calculate_summary_values(pivoted_normal_df, pivoted_total_df)
    if not final_summary_dict:
        logger.error("最终汇总指标计算失败，流程终止。")
        return
        
    logger.info("✅ 最终汇总指标计算成功！")
    
    logger.info("\n--- [步骤 3/3] 展示最终计算结果 ---")
    
    print("\n" + "="*25 + " 最终计算结果 " + "="*25)
    print(json.dumps(final_summary_dict, indent=4, ensure_ascii=False))
    print("="*68)
    
    logger.info("\n========================================")
    logger.info("===         流程执行完毕           ===")
    logger.info("========================================")

if __name__ == '__main__':
    run_audit_report()