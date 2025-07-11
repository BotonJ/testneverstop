from config_loader import ConfigLoader
from data_processor import DataProcessor
from excel_writer import ExcelWriter

# --- 全局配置 ---
MAPPING_FILE = "mapping_file.xlsx"
SOURCE_DATA_FILE = "annual_soce.xlsx"
OUTPUT_REPORT_FILE = "审计报告数据_生成结果.xlsx"

def main():
    """主调度函数，协调所有模块完成报告生成任务。"""
    print("--- 开始执行自动化审计报告生成任务 ---")

    # 1. 加载配置
    config_loader = ConfigLoader(MAPPING_FILE)
    if not config_loader.load_all_sheets():
        print("--- 任务因配置错误而终止 ---")
        return
    
    # 2. 初始化数据处理器
    processor = DataProcessor(SOURCE_DATA_FILE, config_loader.configs)
    
    # 3. 提取并处理数据
    # get_notes_data现在会内部调用解析函数
    notes_data_df = processor.get_notes_data()

    # 如果未能生成任何附注数据，则提前终止
    if notes_data_df.empty:
        print("未能生成任何有效的报表附注数据，任务终止。")
        return

    audit_matters_tables_dict = processor.get_audit_matters_tables() 
    verification_report = processor.run_verification_checks()
    # 4. 生成Excel报告
    writer = ExcelWriter(OUTPUT_REPORT_FILE)    
    # 自动提取年份并生成引言
    audit_year = processor.extract_audit_year()
    if audit_year is None:
        intro_text = "（未能自动获取年份，请手动填写引言）"
    else:
        prev_year = audit_year - 1
        intro_text = (
            f"以下附注项目除特别说明之外，金额单位为人民币元；"
            f"年初是指{audit_year}年1月1日；年末是指{audit_year}年12月31日；"
            f"上年是指{prev_year}年度；本年是指{audit_year}年度。"
        )
    writer.write_notes_sheet(
        sheet_name="报表附注",
        intro_text=intro_text,        
        notes_df=notes_data_df,
        verification_report=verification_report, # <-- 确保这里传递的是这个变量
    )
    
    # 写入第二个Sheet
    if audit_matters_tables_dict:
        writer.write_audit_sheet(
            sheet_name="审计事项说明",
            tables_dict=audit_matters_tables_dict 
        )
    else:
        print("未找到审计事项说明数据，跳过相关Sheet的写入。")  

    # 5. 保存文件
    writer.save()    
    print("--- 任务执行完毕 ---")
if __name__ == '__main__':
    main()