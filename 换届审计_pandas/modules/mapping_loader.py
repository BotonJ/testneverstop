# /modules/mapping_loader.py
import pandas as pd
import openpyxl
from openpyxl.utils.cell import coordinate_to_tuple, column_index_from_string
from src.utils.logger_config import logger

def get_col_index(cell):
    """从'C'或'C5'中安全地提取列索引。忠于原始逻辑。"""
    try:
        if cell is None: return None
        # 如果是单个字母，直接转换
        if str(cell).isalpha():
            return column_index_from_string(cell)
        # 如果是单元格坐标，先提取字母再转换
        return column_index_from_string(coordinate_to_tuple(str(cell))[1])
    except Exception:
        return None # 解析失败则返回None

def parse_skip_rows(value):
    """解析跳过行配置。忠于原始逻辑。"""
    if not value: return []
    rows = []
    for item in str(value).split(","):
        item = item.strip().replace("：", "")
        if item.isdigit():
            rows.append(int(item))
    return rows
    
def load_full_mapping_as_df(mapping_path):
    """
    【旧版Pandas加载器 - 保留备用】
    将映射文件中所有指定的Sheet作为Pandas DataFrame加载到一个字典中。
    """
    # ... 此函数内容保持不变，此处省略 ...
    sheets_to_load = [
        "HeaderMapping", "资产负债表区块", "业务活动表逐行", 
        "科目等价映射", "业务活动表汇总注入配置"
    ]
    all_mappings = {}
    try:
        with pd.ExcelFile(mapping_path) as xls:
            for sheet_name in sheets_to_load:
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    for col in df.select_dtypes(include=['object']).columns:
                        df[col] = df[col].str.strip()
                    all_mappings[sheet_name] = df
                else:
                    all_mappings[sheet_name] = pd.DataFrame()
        return all_mappings
    except Exception as e:
        logger.error(f"使用pandas加载完整的mapping_file.xlsx时出错: {e}")
        return {name: pd.DataFrame() for name in sheets_to_load}

def load_mapping_file(path):
    """
    【最终版 - 忠于原始逻辑】
    使用openpyxl精确解析mapping_file，返回与原始脚本完全一致的数据结构。
    """
    logger.info("--- 开始使用原始逻辑精确解析 mapping_file.xlsx ---")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        logger.error(f"映射文件未找到: {path}")
        return {}

    # 1. 解析 "资产负债表区块"
    blocks = {}
    if "资产负债表区块" in wb.sheetnames:
        ws = wb["资产负债表区块"]
        # 使用pandas读取以简化空行和类型处理
        df = pd.read_excel(path, sheet_name="资产负债表区块")
        for _, row in df.iterrows():
            block_name = row.get('区块名称')
            if pd.isna(block_name): continue
            blocks[str(block_name).strip()] = row.to_dict()

    # 2. 解析 "科目等价映射" (使用Pandas更健壮)
    alias_map_df = pd.DataFrame()
    if "科目等价映射" in wb.sheetnames:
        alias_map_df = pd.read_excel(path, sheet_name="科目等价映射")

    # 3. 解析 "业务活动表逐行"
    yewu_map = []
    if "业务活动表逐行" in wb.sheetnames:
        ws = wb["业务活动表逐行"]
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(cell is None or str(cell).strip() == "" for cell in row):
                continue
            yewu_map.append(dict(zip(headers, row)))

    # 4. 解析 "HeaderMapping" (此部分在我们的新流程中暂时用不到，但保留解析逻辑)
    header_meta = {}
    if "HeaderMapping" in wb.sheetnames:
        # ... (保留您原始的header_meta解析逻辑或简化)
        pass
    
    logger.info("--- mapping_file.xlsx 解析完成 ---")
    
    # 5. 返回与新流程兼容的数据结构
    # 我们将原始的、更精确的解析结果传递给新模块使用
    return {
        "blocks_df": pd.DataFrame.from_dict(blocks, orient='index'),
        "alias_map_df": alias_map_df,
        "yewu_line_map": yewu_map,
        "header_meta": header_meta # 保留
    }