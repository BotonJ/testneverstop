# /main.py

import sys
import os
import json
import logging

# --- [修复] 修正模块导入路径和函数名 ---
# 根据您的项目结构，从 'modules' 导入 'load_mapping_file'
from modules.mapping_loader import load_mapping_file
from src.legacy_runner import run_legacy_extraction
from src.data_processor import pivot_and_clean_data, calculate_summary_values
from src.data_validator import run_all_checks

# 直接配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# --- 确保 'src' 和 'modules' 目录都在Python的搜索路径中 ---
project_root = os.path.dirname(os.path.abspath(__file__))
for folder in ['src', 'modules']:
    path_to_add = os.path.join(project_root, folder)
    if path_to_add not in sys.path:
        sys.path.append(path_to_add)

def run_audit_report():
    logger.info("========================================")
    logger.info("===    自动化审计报告生成流程启动    ===")
    logger.info("========================================")

    # --- 文件路径管理 ---
    source_file = os.path.join(project_root, 'data', 'soce.xlsx')
    mapping_file = os.path.join(project_root, 'data', 'mapping_file.xlsx')
    os.makedirs(os.path.join(project_root, 'output'), exist_ok=True)
    
    # --- [流程一] 加载配置并提取数据 ---
    logger.info("\n--- [流程一] 开始加载配置并提取数据 ---")
    
    # 1a. 加载配置文件，得到配置字典
    full_mapping = load_mapping_file(mapping_file)
    if not full_mapping:
        logger.error("配置文件加载失败，流程终止。")
        return

    # 1b. [修复] 将加载好的配置字典(full_mapping)传递给提取函数
    raw_df = run_legacy_extraction(source_file, full_mapping)
    
    if raw_df is not None and not raw_df.empty:
        logger.info("\n--- [流程二] 开始处理数据并计算核心指标 ---")
        pivoted_normal_df, pivoted_total_df = pivot_and_clean_data(raw_df)
        
        # --- [BUG修复] 传入函数所需的 'raw_df' 参数 ---
        final_summary_dict = calculate_summary_values(pivoted_total_df, raw_df)
        
        logger.info("\n--- [流程三] 开始执行数据交叉复核 (目标A) ---")
        verification_results_list, _ = run_all_checks(raw_df, pivoted_total_df, final_summary_dict, full_mapping)
        
        # --- 在终端打印分析结果 ---
        print("\n" + "="*25 + " Pandas分析结果 " + "="*25)
        # 使用default=str来处理可能无法序列化的数据类型
        print(json.dumps(final_summary_dict, indent=4, ensure_ascii=False, default=str)) 
        print("\n" + "="*27 + " 复核报告摘要 " + "="*27)
        if verification_results_list:
            for line in verification_results_list:
                print(line)
        else:
            print("✅ 复核机制未发现问题。")
        print("="*68)
    else:
        logger.error("数据提取失败，无法进行后续分析。")

    # --- 目标B已暂停 ---
    # logger.info("\n--- [流程四] 开始生成格式化的'新soce'报告 ---")

    logger.info("\n========================================")
    logger.info("===           流程执行完毕           ===")
    logger.info("========================================")

if __name__ == '__main__':
    run_audit_report()