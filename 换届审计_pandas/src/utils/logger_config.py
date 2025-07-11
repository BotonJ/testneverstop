# /src/utils/logger_config.py

import logging
import os

def setup_logger():
    """
    设置一个全局的、双输出的日志记录器。
    """
    # 获取项目的根目录
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    log_dir = os.path.join(project_root, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_filepath = os.path.join(log_dir, 'audit_run.log')

    # 1. 获取一个日志记录器实例
    # 使用一个固定的名字，确保在项目各处获取的是同一个logger实例
    logger = logging.getLogger("AuditReportLogger")
    logger.setLevel(logging.DEBUG)  # 设置logger的最低处理级别为DEBUG

    # 防止重复添加handler
    if logger.hasHandlers():
        logger.handlers.clear()

    # 2. 创建一个用于输出到文件的Handler
    # 这个handler会将所有DEBUG及以上级别的日志都写入文件
    file_handler = logging.FileHandler(log_filepath, mode='w', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(module)s.%(funcName)s:%(lineno)d - %(message)s"
    )
    file_handler.setFormatter(file_formatter)

    # 3. 创建一个用于输出到控制台的Handler
    # 这个handler只显示INFO及以上级别的日志，保持控制台输出简洁
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s"
    )
    console_handler.setFormatter(console_formatter)

    # 4. 将两个Handler添加到logger中
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# 执行一次设置，并导出一个可以直接使用的logger实例
logger = setup_logger()