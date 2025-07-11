# /modules/mapping_loader.py
import pandas as pd
from src.utils.logger_config import logger

def load_mapping_file(mapping_path):
    """
    加载并解析最核心的映射文件配置。
    (这个旧函数我们暂时保留，以备不时之需)
    """
    # ... 此函数内容保持不变 ...
    try:
        mapping_wb = pd.ExcelFile(mapping_path)
        header_df = pd.read_excel(mapping_wb, sheet_name="HeaderMapping")
        blocks_df = pd.read_excel(mapping_wb, sheet_name="资产负债表区块")
        alias_df = pd.read_excel(mapping_wb, sheet_name="科目等价映射")

        alias_map = {
            str(row["标准科目名"]).strip(): [
                str(alias).strip()
                for alias in row.to_list()[1:]
                if pd.notna(alias)
            ]
            for _, row in alias_df.iterrows()
        }

        return {
            "header": header_df.set_index("配置项")["配置值"].to_dict(),
            "blocks": blocks_df.to_dict("records"),
            "subject_alias_map": alias_map,
        }
    except Exception as e:
        logger.error(f"加载 mapping_file.xlsx 时出错: {e}")
        return None

def load_full_mapping_as_df(mapping_path):
    """
    【新函数】
    将映射文件中所有指定的Sheet作为Pandas DataFrame加载到一个字典中。
    """
    sheets_to_load = [
        "HeaderMapping",
        "资产负债表区块",
        "业务活动表逐行",
        "科目等价映射",
        "业务活动表汇总注入配置" # 即使暂时不用，也预先加载
    ]
    
    all_mappings = {}
    try:
        with pd.ExcelFile(mapping_path) as xls:
            for sheet_name in sheets_to_load:
                if sheet_name in xls.sheet_names:
                    all_mappings[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                else:
                    logger.warning(f"在 {mapping_path} 中未找到名为 '{sheet_name}' 的Sheet，将跳过。")
                    all_mappings[sheet_name] = pd.DataFrame() # 未找到则返回空DataFrame
        return all_mappings
    except Exception as e:
        logger.error(f"使用pandas加载完整的mapping_file.xlsx时出错: {e}")
        return {name: pd.DataFrame() for name in sheets_to_load} # 出错时返回空的DataFrame字典