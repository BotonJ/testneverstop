# /modules/utils.py

def normalize_name(name):
    """
    【经典代码 - 核心清洗函数】
    一个健壮的函数，用于清洗和标准化科目名称。
    它处理各种空格、特殊字符和大小写问题。
    """
    if not isinstance(name, str):
        return ""
    
    # 替换全角字符和常见特殊符号为空格
    replacements = {
        '（': '(', '）': ')', '：': ':', '　': ' ', '－': '-'
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
        
    # 移除所有内部和外部的空格
    name = "".join(name.split())
    
    # 可选：统一转为小写以便不区分大小写比较
    # name = name.lower()
    
    return name