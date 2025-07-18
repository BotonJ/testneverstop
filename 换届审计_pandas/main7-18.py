# /main.py

import sys
import os
import logging
import pandas as pd

# --- [V5.0 融合版] ---
from modules.mapping_loader import load_mapping_file
from src.legacy_runner import run_legacy_extraction
from src.data_processor import calculate_summary_values
from src.report_generator import generate_master_report
from modules.utils import normalize_name

# --- 日志配置 ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# --- 路径管理 ---
project_root = os.path.dirname(os.path.abspath(__file__))
# (sys.path.append...)

def run_audit_report():
    logger.info("========================================")
    logger.info("===    自动化审计报告生成流程启动 (V5.0 融合版)    ===")
    logger.info("========================================")

    # --- 1. 路径和配置管理 ---
    source_file = os.path.join(project_root, 'data', 'soce.xlsx')
    mapping_file_path = os.path.join(project_root, 'data', 'mapping_file.xlsx')
    template_file = os.path.join(project_root, 'data', 't.xlsx')
    output_dir = os.path.join(project_root, 'output')
    os.makedirs(output_dir, exist_ok=True)
    final_report_file = os.path.join(output_dir, '最终审计报告.xlsx')

    # --- 2. 加载配置 ---
    logger.info("\n--- 步骤 1/4: 加载配置文件 ---")
    mapping_configs = load_mapping_file(mapping_file_path)
    if not mapping_configs: return
    # 为方便后续使用，添加几个关键路径和字典到配置中
    mapping_configs['template_path'] = template_file
    alias_df = mapping_configs.get('科目等价映射', pd.DataFrame())

    alias_dict_builder = {}
    if not alias_df.empty:
        # 将标准名和它自身先映射起来
        for std_name in alias_df['标准科目名'].unique():
            if pd.notna(std_name):
                norm_std = normalize_name(std_name)
                alias_dict_builder[norm_std] = norm_std
        
        # 遍历所有等价科目列来构建别名映射
        for _, row in alias_df.iterrows():
            std_name = row.get('标准科目名')
            if pd.isna(std_name): continue
            norm_std = normalize_name(std_name)

            for col in alias_df.columns:
                if '等价科目名' in col and pd.notna(row[col]):
                    # 关键修复：先用str()转换为字符串，再split
                    aliases_str = str(row[col])
                    for alias in aliases_str.split(','):
                        # 清理别名并检查有效性
                        cleaned_alias = alias.strip()
                        if cleaned_alias: # 确保不是空字符串
                            alias_dict_builder[normalize_name(cleaned_alias)] = norm_std
                            
    mapping_configs['alias_dict'] = alias_dict_builder
    # --- 3. 提取数据并生成内存预制件 ---
    logger.info("\n--- 步骤 2/4: 提取数据并生成内存预制件 ---")
    raw_df, prebuilt_wb = run_legacy_extraction(source_file, mapping_configs)
    if raw_df is None or raw_df.empty: return

    # --- 4. 计算核心汇总值 ---
    logger.info("\n--- 步骤 3/4: 计算核心汇总值 ---")
    summary_values = calculate_summary_values(raw_df)
    
    # --- 5. 生成最终报告 ---
    logger.info("\n--- 步骤 4/4: 生成统一的主审计报告 ---")
    generate_master_report(
        prebuilt_wb=prebuilt_wb,
        mapping_configs=mapping_configs,
        output_path=final_report_file,
        summary_values=summary_values
    )

    logger.info("\n========================================")
    logger.info("===           流程执行完毕           ===")
    logger.info("========================================")

if __name__ == '__main__':
    run_audit_report()