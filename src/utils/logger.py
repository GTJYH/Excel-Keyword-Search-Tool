#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel关键字搜索工具日志工具模块
"""

import logging
import os
from datetime import datetime

def setup_logging():
    """设置日志配置"""
    # 如果日志目录不存在则创建
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # 创建带时间戳的日志文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"excel_search_{timestamp}.log")
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def get_logger(name):
    """获取指定名称的日志记录器实例"""
    return logging.getLogger(name)