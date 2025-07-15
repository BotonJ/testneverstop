import pandas as pd
import logging
import os

# 导入我们所有的自定义模块
from src import mapping_loader
from src import legacy_runner
from src import data_processor
from src import data_validator
# from src import classic_report_generator # 遵照指示，暂时禁用目标B

# 配置日志记录
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def main():
    """
    自动化审计报告生成的主执行函数。
    """
    logging.info("========================================")
    logging.info("===    自动化审计报告生成流程启动    ===")
    logging.info("========================================")

    # --- [配置] 定义核心文件路径 ---
    # 使用os.path.join确保路径在不同操作系统下都能正常工作
    base_path = 'D:\\python脚本合集\\审计自动化\\my_github_repos\\换届审计_pandas'
    data_path = os.path.join(base_path, 'data')
    output_path = os.path.join(base_path, 'output')
    
    source_file = os.path.join(data_path, "soce.xlsx")
    mapping_file = os.path.join(data_path, "mapping_file.xlsx")
    # template_file = os.path.join(data_path, "t.xlsx") # 暂时禁用

    # 确保输出目录存在
    os.makedirs(output_path, exist_ok=True)
    
    # --- [流程一] 加载与提取 ---
    logging.info("\n--- [步骤 1/4] 开始加载配置文件并提取原始数据 ---")
    full_mapping = mapping_loader.load_mapping_files(mapping_file)
    raw_df = legacy_runner.run_legacy_extraction(source_file, full_mapping)

    # --- [流程二] 处理与计算 ---
    if raw_df is not None and not raw_df.empty:
        logging.info("\n--- [步骤 2/4] 开始处理数据并计算核心指标 ---")
        
        # 2a. 对原始数据进行透视，拆分为普通科目表和合计科目表
        pivoted_normal_df, pivoted_total_df = data_processor.pivot_and_clean_data(raw_df)
        
        # 2b. 基于合计科目表，计算最终的汇总指标字典
        final_summary_dict = data_processor.calculate_summary_values(pivoted_total_df)

        logging.info("Pandas分析结果:\n" + pd.Series(final_summary_dict).to_string())

        # --- [流程三] 复核与验证 (目标A) ---
        logging.info("\n--- [步骤 3/4] 开始执行数据交叉复核 ---")
        verification_results_list, _ = data_validator.run_all_checks(
            raw_df, 
            pivoted_total_df, 
            final_summary_dict, 
            full_mapping
        )
        
        # 打印复核结果摘要
        print("\n=========================== 复核报告摘要 ===========================")
        if verification_results_list:
            for result in verification_results_list:
                print(result)
        else:
            print("✅ 未发现复核问题。")
        print("====================================================================")

        # --- [流程四] 生成报告 (目标B - 已暂停) ---
        # logging.info("\n--- [步骤 4/4] 开始生成最终格式化报告 ---")
        # classic_report_generator.run_classic_report_generation(
        #     pivoted_total_df,
        #     final_summary_dict,
        #     template_file,
        #     output_path,
        #     full_mapping
        # )

    else:
        logging.error("未能从源文件提取任何数据，流程终止。")

    logging.info("\n========================================")
    logging.hinfo("===           流程执行完毕           ===")
    logging.info("========================================")


if __name__ == '__main__':
    main()