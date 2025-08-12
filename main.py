#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel关键字搜索工具
一个专业的工具，用于在目录中搜索Excel文件中的关键字。
"""

import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from src.main_window import MainWindow
from src.utils.logger import setup_logging

def main():
    """主应用程序入口点"""
    # 设置日志
    setup_logging()
    
    # 创建Qt应用程序
    app = QApplication(sys.argv)
    app.setApplicationName("Excel关键字搜索工具")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("专业工具")
    
    # 设置应用程序样式 - 使用自定义明亮主题
    # 移除深色主题，使用MainWindow中的自定义样式
    
    # 创建并显示主窗口
    window = MainWindow()
    window.show()
    
    # 启动事件循环
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
