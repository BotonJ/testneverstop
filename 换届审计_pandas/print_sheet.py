# 确保您已经在项目根目录创建了print_sheet.py并粘贴了以上代码。
# 打开您的终端（命令行工具），并确保当前路径位于项目根目录下。
# 要打印科目等价映射的全部内容，请输入并执行以下命令：
#python print_sheet.py "C:\审计自动化\my_github_repos\换届审计_pandas\mapping_file.xlsx" "  科目等价映射""，如果您的Sheet名包含空格，请务必用双引号""把它括起来。
import pandas as pd
import sys
import os

def print_sheet_content(file_path, sheet_name):
    """
    读取并打印指定Excel文件中特定Sheet的全部内容。
    """
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 '{file_path}'")
        return

    try:
        # 使用pandas读取指定的sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        print("-" * 50)
        print(f"文件: '{file_path}'")
        print(f"Sheet页: '{sheet_name}'")
        print("-" * 50)
        
        if df.empty:
            print("这个Sheet页是空的。")
        else:
            # 使用to_string()来确保所有行和列都被完整打印
            print(df.to_string())
            
        print("-" * 50)

    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        print("请检查：")
        print(f"1. 文件路径 '{file_path}' 是否正确。")
        print(f"2. Sheet页名称 '{sheet_name}' 是否存在于文件中，且没有拼写错误。")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使用方法: python print_sheet.py <文件路径> \"<Sheet页名称>\"")
        print("示例: python print_sheet.py \"data/mapping_file.xlsx\" \"科目等价映射\"")
        sys.exit(1)

    file_to_read = sys.argv[1]
    sheet_to_read = sys.argv[2]
    
    print_sheet_content(file_to_read, sheet_to_read)

