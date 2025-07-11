import pandas as pd

# --- 配置 ---
# 请确保这个文件名与您的 mapping 文件名完全一致
MAPPING_FILE = r"D:\python脚本合集\审计自动化\my_github_repos\换届审计_pandas\data\mapping_file.xlsx"

def print_mapping_headers(filepath: str):
    """
    读取指定的Excel文件，并打印出其中每一个Sheet页的表头（列名）。
    
    :param filepath: Excel文件的路径。
    """
    print(f"--- 开始读取文件: '{filepath}' ---")
    
    try:
        # 使用 pandas.ExcelFile 可以高效地获取所有sheet的名称，而无需一次性加载所有数据
        xls = pd.ExcelFile(filepath)
        sheet_names = xls.sheet_names
        
        print(f"✅ 文件读取成功，共发现 {len(sheet_names)} 个Sheet页。")
        print("="*50)
        
        # 遍历每一个Sheet页
        for sheet_name in sheet_names:
            try:
                # 读取单个sheet来获取其表头
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # 获取表头列表
                headers = df.columns.tolist()
                
                # 打印结果
                print(f"Sheet页名称: '{sheet_name}'")
                print(f"表头 (列名): {headers}")
                print("-"*50)
                
            except Exception as e:
                print(f"⚠️ 读取Sheet页 '{sheet_name}' 时发生错误: {e}")
                print("-"*50)

    except FileNotFoundError:
        print(f"❌ 错误：无法找到文件 '{filepath}'。请确保文件名正确，并且文件与脚本在同一目录下。")
    except Exception as e:
        print(f"❌ 错误：读取Excel文件时发生未知错误: {e}")

if __name__ == '__main__':
    print_mapping_headers(MAPPING_FILE)