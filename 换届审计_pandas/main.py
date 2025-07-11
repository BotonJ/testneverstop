# /main.py

import json
import os
import sys
# --- 核心设置：将src目录添加到Python的模块搜索路径 ---
# 这样main.py就能找到位于src文件夹下的模块了
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.append(PROJECT_ROOT)

# --- 模块导入 ---
# 从我们重构好的模块中导入核心功能
from src.utils.logger_config import logger
from src.legacy_runner import run_legacy_extraction
from src.data_processor import pivot_and_clean_data, calculate_summary_values

def run_audit_report():
    """
    【主流程函数】
     orchestrates the entire audit report generation process.
    """
    logger.info("========================================")
    logger.info("===    自动化审计报告生成流程启动    ===")
    logger.info("========================================")

    # --- 1. 定义文件路径 ---
    # 使用os.path.join确保跨平台兼容性
    project_root = os.path.dirname(os.path.abspath(__file__))
    source_file = os.path.join(project_root, 'data', 'soce.xlsx')
    mapping_file = os.path.join(project_root, 'data', 'mapping_file.xlsx')
    
    logger.info(f"源文件路径: {source_file}")
    logger.info(f"映射文件路径: {mapping_file}")

    # --- 2. 步骤一：提取原始数据 ---
    # 调用legacy_runner，获取原始的长格式DataFrame
    logger.info("\n--- [步骤 1/3] 执行数据提取 ---")
    raw_df = run_legacy_extraction(source_file, mapping_file)

    if raw_df is None:
        logger.error("数据提取失败，流程终止。请检查legacy_runner的日志输出。")
        return

    logger.info("✅ 数据提取成功！原始DataFrame已加载到内存。")
    # print("\n原始DataFrame预览:\n", raw_df.head()) # 取消注释以查看详细输出

    # --- 3. 步骤二：数据处理与计算 ---
    # 调用data_processor，进行数据透视和指标计算
    logger.info("\n--- [步骤 2/3] 执行数据处理与计算 ---")
    
    # a. 数据透视
    pivoted_df = pivot_and_clean_data(raw_df)
    if pivoted_df is None or pivoted_df.empty:
        logger.error("数据透视失败，流程终止。")
        return
    logger.info("✅ 数据透视与清理成功！")
    # print("\n透视后DataFrame预览:\n", pivoted_df.head()) # 取消注释以查看详细输出
        
    # b. 计算最终指标
    final_summary_dict = calculate_summary_values(pivoted_df)
    if not final_summary_dict:
        logger.error("最终汇总指标计算失败，流程终止。")
        return
        
    logger.info("✅ 最终汇总指标计算成功！")
    
    # --- 4. 步骤三：展示最终结果 ---
    # 目前我们先将结果打印出来，后续再加入报告生成模块
    logger.info("\n--- [步骤 3/3] 展示最终计算结果 ---")
    
    print("\n" + "="*25 + " 最终计算结果 " + "="*25)
    # 使用json.dumps美化字典的打印输出
    print(json.dumps(final_summary_dict, indent=4, ensure_ascii=False))
    print("="*68)
    
    logger.info("\n========================================")
    logger.info("===         流程执行完毕           ===")
    logger.info("========================================")


# --- 脚本执行入口 ---
if __name__ == '__main__':
    run_audit_report()