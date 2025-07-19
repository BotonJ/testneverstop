import pandas as pd
import sys

# --- 配置 ---
# 请确保这个文件名与您的 mapping 文件名完全一致
MAPPING_FILE = r"D:\python脚本合集\审计自动化\my_github_repos\换届审计_pandas\data\mapping_file.xlsx"

def print_full_sheet_content(filepath: str, sheet_name_to_print: str):
    """
    读取指定的Excel文件，并将某一个特定Sheet页的全部内容打印到控制台，
    不省略任何行或列。
    
    :param filepath: Excel文件的路径。
    :param sheet_name_to_print: 需要打印的Sheet页的名称。
    """
    print(f"--- 准备打印文件 '{filepath}' 中 '{sheet_name_to_print}' 的全部内容 ---")
    
    try:
        # 使用 pandas 读取指定的Sheet页
        # 我们在这里不需要一次性读取所有sheet，只读需要的即可
        df = pd.read_excel(filepath, sheet_name=sheet_name_to_print)
        
        # --- 核心设置：保证Pandas打印全部内容 ---
        # 设置打印时显示所有列（不会用...省略）
        pd.set_option('display.max_columns', None)
        # 设置打印时显示所有行
        pd.set_option('display.max_rows', None)
        # 设置每列的最大宽度，以便显示长文本
        pd.set_option('display.max_colwidth', None)
        # 设置打印宽度，以利用整个控制台宽度
        pd.set_option('display.width', 1000)
        
        print(f"✅ Sheet页 '{sheet_name_to_print}' 读取成功，内容如下：\n")
        
        # 打印整个 DataFrame
        print(df)
        
        print(f"\n--- '{sheet_name_to_print}' 内容打印完毕 ---")

    except FileNotFoundError:
        print(f"❌ 错误：无法找到文件 '{filepath}'。请确保文件名正确。")
    except ValueError:
        # 当pd.read_excel找不到指定的sheet时，会抛出ValueError
        print(f"❌ 错误：在文件 '{filepath}' 中找不到名为 '{sheet_name_to_print}' 的Sheet页。")
        # 打印出所有可用的sheet名，方便用户检查
        try:
            xls = pd.ExcelFile(filepath)
            print(f"  文件中可用的Sheet页有: {xls.sheet_names}")
        except Exception:
            pass # 如果连文件都打不开，就忽略
    except Exception as e:
        print(f"❌ 错误：发生未知错误: {e}")

if __name__ == '__main__':
    # 从命令行获取要打印的sheet名
    # 如果用户没有提供，就打印提示信息
    if len(sys.argv) > 1:
        sheet_to_display = sys.argv[1]
        print_full_sheet_content(MAPPING_FILE, sheet_to_display)
    else:
        print("请在运行时提供要打印的Sheet页名称。")
        print("用法示例: python print_sheet_content.py 资产负债表区块")
        # 作为一个友好的备用方案，我们可以打印出所有sheet的表头
        try:
            print("\n--- 文件中所有Sheet页的表头如下 ---")
            xls = pd.ExcelFile(MAPPING_FILE)
            for sheet in xls.sheet_names:
                df_head = pd.read_excel(xls, sheet_name=sheet)
                print(f"\nSheet: '{sheet}'\n表头: {df_head.columns.tolist()}")
        except FileNotFoundError:
            print(f"无法找到 '{MAPPING_FILE}' 文件。")
        except Exception as e:
            print(f"读取文件时出错: {e}")