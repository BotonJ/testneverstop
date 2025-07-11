import pandas as pd
from typing import Dict, List

class ConfigLoader:
    """
    负责加载和验证 mapping_file.xlsx 的配置。
    它将所有定义的Sheet页加载到内存中，并提供一个统一的访问接口。
    """
    # 定义建议存在的Sheet页名称，脚本会尝试加载它们
    SHEET_NAMES: List[str] = [
        "资产负债表区块",
        "业务活动表逐行",
        "科目等价映射",
        "inj1",
        "text_mapping"
    ]

    def __init__(self, mapping_filepath: str):
        self.filepath = mapping_filepath
        self.configs: Dict[str, pd.DataFrame] = {}
        print(f"初始化配置加载器，目标文件: {self.filepath}")

    def load_all_sheets(self) -> bool:
        """
        加载所有定义的Sheet页到 self.configs 字典中。
        """
        try:
            print("开始加载 mapping_file.xlsx...")
            with pd.ExcelFile(self.filepath) as xls:
                # 循环加载所有定义的Sheet页
                for sheet_name in self.SHEET_NAMES:
                    if sheet_name in xls.sheet_names:
                        self.configs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                        print(f"  ✅ 成功加载Sheet页: '{sheet_name}'")
                    else:
                        print(f"  ⚠️ 信息：在文件中未找到可选的Sheet页: '{sheet_name}'，已跳过。")
            
            print("所有可用的配置Sheet页已加载。")
            return True

        except FileNotFoundError:
            print(f"❌ 错误：无法找到配置文件，请检查路径是否正确: {self.filepath}")
            return False
        except Exception as e:
            print(f"❌ 错误：在加载配置文件时发生未知错误: {e}")
            return False

    def get_config_df(self, sheet_name: str) -> pd.DataFrame:
        """
        提供一个安全的接口来获取已加载的配置DataFrame。
        """
        if sheet_name not in self.configs:
            # 返回一个空的DataFrame而不是报错，让调用方可以安全处理
            return pd.DataFrame()
        
        return self.configs[sheet_name]