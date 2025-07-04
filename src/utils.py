# src/utils.py

import sys
import os
from pathlib import Path

def get_base_path():
    """
    获取程序运行的基础路径（即 .exe 文件所在的目录）。
    这段代码可以被任何一个脚本安全地调用。
    """
    if getattr(sys, 'frozen', False):
        # 打包后运行
        return Path(os.path.dirname(sys.executable))
    else:
        # 在开发环境中直接运行 .py
        return Path(os.path.dirname(os.path.abspath(__file__))).parent