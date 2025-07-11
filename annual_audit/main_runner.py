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
    
    # 2. 加载并处理数据
    processor = DataProcessor(SOURCE_DATA_FILE, config_loader)
    if not processor.load_source_data():
        print("--- 任务因源数据错误而终止 ---")
        return
        
    # 获取处理好的数据
    notes_data_df = processor.get_notes_data()
    audit_data_dict = processor.get_audit_matters_data()

    # 3. 生成Excel报告
    writer = ExcelWriter(OUTPUT_REPORT_FILE)
    
    # 准备报表附注的引言文本
    text_config = config_loader.get_config_df('text_mapping')
    audit_year = text_config.loc[text_config['item_key'] == 'audit_period_text', 'value_source'].iloc[0]
    intro_text = f"以下附注项目除特别说明之外，金额单位为人民币元；年初是指{audit_year}年1月1日；年末是指{audit_year}年12月31日；上年是指{audit_year}年度；本年是指{audit_year}年度。"
    
    # 写入第一个Sheet
    writer.write_notes_sheet(
        sheet_name="报表附注",
        intro_text=intro_text,
        notes_df=notes_data_df
    )
    
    # 写入第二个Sheet
    writer.write_audit_sheet(
        sheet_name="审计事项说明",
        audit_data=audit_data_dict
    )

    # 4. 保存文件
    writer.save()
    
    print("--- 任务执行完毕 ---")


if __name__ == '__main__':
    main()