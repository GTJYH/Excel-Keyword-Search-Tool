#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件表格组件模块
"""

import os
import subprocess
import platform
from PyQt6.QtWidgets import (
    QTableWidget, QTableWidgetItem, QHeaderView, QMenu, QMessageBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QCursor
from ..utils.logger import get_logger
from ..utils.i18n import get_text

logger = get_logger(__name__)

class FileTableWidget(QTableWidget):
    """文件结果表格组件"""
    
    # 信号定义
    file_double_clicked = pyqtSignal(str)  # 双击文件时发出信号
    
    def __init__(self):
        super().__init__()
        self.file_data = []  # 存储文件数据
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        # 设置表格属性
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels([
            get_text("File Name"),
            get_text("Path"),
            get_text("Size"),
            get_text("Modified"),
            get_text("Keyword Matches")
        ])
        
        # 设置表格样式
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.setSortingEnabled(True)
        
        # 隐藏行号
        self.verticalHeader().setVisible(False)
        
        # 设置网格
        self.setShowGrid(True)
        self.setGridStyle(Qt.PenStyle.SolidLine)
        
        # 禁用焦点框
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        

        
        # 设置列宽
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # 文件名 - 固定宽度
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)  # 路径 - 自适应
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)  # 大小 - 固定宽度
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)  # 修改时间 - 固定宽度
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)  # 匹配数 - 固定宽度
        
        # 设置固定列宽
        self.setColumnWidth(0, 200)  # 文件名列宽
        self.setColumnWidth(2, 80)   # 大小列宽
        self.setColumnWidth(3, 120)  # 修改时间列宽
        self.setColumnWidth(4, 100)  # 匹配数列宽
        
        # 设置右键菜单
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
        # 连接双击信号
        self.itemDoubleClicked.connect(self.on_item_double_clicked)
        
    def add_file(self, file_info):
        """添加文件到表格"""
        row = self.rowCount()
        self.insertRow(row)
        logger.info(f"添加文件到表格第{row}行: {file_info['name']}")
        
        # 创建表格项
        name_item = QTableWidgetItem(file_info['name'])
        path_item = QTableWidgetItem(file_info['path'])
        size_item = QTableWidgetItem(file_info['size'])
        modified_item = QTableWidgetItem(file_info['modified'])
        matches_item = QTableWidgetItem(str(file_info['matches']))
        

        
        # 设置工具提示
        name_item.setToolTip(file_info['name'])
        path_item.setToolTip(file_info['path'])
        
        # 设置项目不可编辑
        name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        path_item.setFlags(path_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        size_item.setFlags(size_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        modified_item.setFlags(modified_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        matches_item.setFlags(matches_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        
        # 添加到表格
        self.setItem(row, 0, name_item)
        self.setItem(row, 1, path_item)
        self.setItem(row, 2, size_item)
        self.setItem(row, 3, modified_item)
        self.setItem(row, 4, matches_item)
        
        # 存储文件数据
        self.file_data.append(file_info)
        
    def get_file_info(self, row):
        """获取指定行的文件信息"""
        if 0 <= row < len(self.file_data):
            return self.file_data[row]
        return None
        
    def clear(self):
        """清空表格"""
        # 只清除行数据，保留表头
        self.setRowCount(0)
        self.file_data.clear()
        
    def on_item_double_clicked(self, item):
        """处理项目双击事件"""
        row = item.row()
        file_info = self.get_file_info(row)
        if file_info:
            self.file_double_clicked.emit(file_info['path'])
            self.open_file(file_info['path'])
            
    def show_context_menu(self, position):
        """显示右键菜单"""
        row = self.rowAt(position.y())
        if row >= 0:
            file_info = self.get_file_info(row)
            if file_info:
                self.show_file_context_menu(file_info, position)
                
    def show_file_context_menu(self, file_info, position):
        """显示文件右键菜单"""
        menu = QMenu(self)
        
        # 打开文件动作
        open_action = QAction(get_text("Open File"), self)
        open_action.triggered.connect(lambda: self.open_file(file_info['path']))
        menu.addAction(open_action)
        
        # 打开文件夹动作
        open_folder_action = QAction(get_text("Open Folder"), self)
        open_folder_action.triggered.connect(lambda: self.open_folder(file_info['path']))
        menu.addAction(open_folder_action)
        
        menu.addSeparator()
        
        # 复制路径动作
        copy_path_action = QAction(get_text("Copy Path"), self)
        copy_path_action.triggered.connect(lambda: self.copy_path(file_info['path']))
        menu.addAction(copy_path_action)
        
        # 复制文件名动作
        copy_name_action = QAction(get_text("Copy File Name"), self)
        copy_name_action.triggered.connect(lambda: self.copy_name(file_info['name']))
        menu.addAction(copy_name_action)
        
        # 显示菜单
        menu.exec(QCursor.pos())
        
    def open_file(self, file_path):
        """打开文件"""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
                
            logger.info(f"已打开文件: {file_path}")
            
        except Exception as e:
            logger.error(f"打开文件失败 {file_path}: {str(e)}")
            QMessageBox.warning(self, get_text("Warning"), 
                              f"{get_text('Failed to open file')}: {str(e)}")
                              
    def open_folder(self, file_path):
        """打开文件所在文件夹"""
        try:
            folder_path = os.path.dirname(file_path)
            
            if platform.system() == "Windows":
                subprocess.run(["explorer", folder_path])
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
                
            logger.info(f"已打开文件夹: {folder_path}")
            
        except Exception as e:
            logger.error(f"打开文件夹失败 {folder_path}: {str(e)}")
            QMessageBox.warning(self, get_text("Warning"), 
                              f"{get_text('Failed to open folder')}: {str(e)}")
                              
    def copy_path(self, file_path):
        """复制文件路径到剪贴板"""
        try:
            clipboard = self.window().window().clipboard()
            clipboard.setText(file_path)
            logger.info(f"已复制路径到剪贴板: {file_path}")
        except Exception as e:
            logger.error(f"复制路径失败: {str(e)}")
            
    def copy_name(self, file_name):
        """复制文件名到剪贴板"""
        try:
            clipboard = self.window().window().clipboard()
            clipboard.setText(file_name)
            logger.info(f"已复制文件名到剪贴板: {file_name}")
        except Exception as e:
            logger.error(f"复制文件名失败: {str(e)}")
            
    def get_selected_files(self):
        """获取选中的文件列表"""
        selected_files = []
        for row in range(self.rowCount()):
            if self.item(row, 0).isSelected():
                file_info = self.get_file_info(row)
                if file_info:
                    selected_files.append(file_info)
        return selected_files
        
    def select_all(self):
        """全选"""
        self.selectAll()
        
    def clear_selection(self):
        """清除选择"""
        self.clearSelection()
        
    def sort_by_column(self, column, order):
        """按列排序"""
        self.sortItems(column, order)
