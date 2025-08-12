#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel关键字搜索工具主窗口模块
"""

import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLineEdit, QLabel, QFileDialog, QTableWidget,
    QTableWidgetItem, QProgressBar, QTextEdit, QSplitter,
    QGroupBox, QCheckBox, QComboBox, QMessageBox, QHeaderView,
    QFrame, QSizePolicy, QApplication
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSettings
from PyQt6.QtGui import QFont, QIcon, QAction, QKeySequence
from .search_engine import SearchEngine
from .components.file_table import FileTableWidget
from .utils.logger import get_logger
from .utils.i18n import get_text, set_language, register_language_change_callback, unregister_language_change_callback

logger = get_logger(__name__)

class MainWindow(QMainWindow):
    """主应用程序窗口"""
    
    def __init__(self):
        super().__init__()
        self.search_engine = None
        self.search_thread = None
        self.settings = QSettings('ProfessionalTools', 'ExcelKeywordSearch')
        
        self.init_ui()
        self.setup_connections()
        self.load_settings()
        
        # 注册语言切换回调
        register_language_change_callback(self.update_ui_language)
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle(get_text("Excel Keyword Search Tool"))
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(16)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 创建搜索面板
        self.create_search_panel(main_layout)
        
        # 创建分割器用于结果和详情
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # 创建结果面板
        self.create_results_panel(splitter)
        
        # 创建详情面板
        self.create_details_panel(splitter)
        
        # 设置分割器比例
        splitter.setSizes([700, 300])
        
        # 创建状态栏
        self.create_status_bar()
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 应用样式
        self.apply_styles()
        
    def create_search_panel(self, parent_layout):
        """创建搜索配置面板"""
        search_group = QGroupBox()
        search_group.setObjectName("search_config_group")
        search_layout = QGridLayout(search_group)
        
        # 目录选择
        self.dir_label = QLabel(get_text("Search Directory:"))
        self.dir_combo = QComboBox()
        self.dir_combo.setEditable(True)
        self.dir_combo.setPlaceholderText(get_text("Select directory to search..."))
        self.dir_combo.setMinimumWidth(300)  # 设置最小宽度
        self.browse_btn = QPushButton(get_text("Browse"))
        self.browse_btn.setFixedWidth(80)  # 固定浏览按钮宽度
        
        # 关键字输入
        self.keyword_label = QLabel(get_text("Keyword:"))
        self.keyword_edit = QLineEdit()
        self.keyword_edit.setPlaceholderText(get_text("Enter keyword to search..."))
        
        # 搜索选项
        self.case_sensitive_cb = QCheckBox(get_text("Case Sensitive"))
        self.whole_word_cb = QCheckBox(get_text("Whole Word"))
        self.include_subdirs_cb = QCheckBox(get_text("Include Subdirectories"))
        self.complete_search_cb = QCheckBox(get_text("Complete Search"))
        self.complete_search_cb.setChecked(True)  # 默认开启完整搜索
        self.complete_search_cb.setToolTip(get_text("Complete Search Tooltip"))
        self.include_subdirs_cb.setChecked(True)
        
        # 文件类型过滤器
        self.file_type_label = QLabel(get_text("File Types:"))
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems([
            get_text("All Excel Files (.xlsx, .xls)"),
            get_text("Excel 2007+ (.xlsx)"),
            get_text("Excel 97-2003 (.xls)"),
            get_text("All Files")
        ])
        
        # 搜索按钮
        self.search_btn = QPushButton(get_text("Search"))
        self.search_btn.setMinimumHeight(44)
        
        # 布局
        search_layout.setSpacing(16)
        search_layout.setContentsMargins(20, 20, 20, 20)
        
        search_layout.addWidget(self.dir_label, 0, 0)
        search_layout.addWidget(self.dir_combo, 0, 1)
        search_layout.addWidget(self.browse_btn, 0, 2)
        
        search_layout.addWidget(self.keyword_label, 1, 0)
        search_layout.addWidget(self.keyword_edit, 1, 1, 1, 2)
        
        search_layout.addWidget(self.case_sensitive_cb, 2, 0)
        search_layout.addWidget(self.whole_word_cb, 2, 1)
        search_layout.addWidget(self.include_subdirs_cb, 2, 2)
        
        search_layout.addWidget(self.complete_search_cb, 3, 0)
        
        search_layout.addWidget(self.file_type_label, 4, 0)
        search_layout.addWidget(self.file_type_combo, 4, 1, 1, 2)
        
        search_layout.addWidget(self.search_btn, 5, 0, 1, 3)
        
        parent_layout.addWidget(search_group)
        
    def create_results_panel(self, splitter):
        """创建结果显示面板"""
        results_group = QGroupBox()
        results_group.setObjectName("search_results_group")
        results_layout = QVBoxLayout(results_group)
        results_layout.setSpacing(12)
        results_layout.setContentsMargins(16, 16, 16, 16)
        
        # 结果表格
        self.results_table = FileTableWidget()
        results_layout.addWidget(self.results_table)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        results_layout.addWidget(self.progress_bar)
        
        splitter.addWidget(results_group)
        
    def create_details_panel(self, splitter):
        """创建详情面板"""
        details_group = QGroupBox()
        details_group.setObjectName("file_details_group")
        details_layout = QVBoxLayout(details_group)
        details_layout.setSpacing(12)
        details_layout.setContentsMargins(16, 16, 16, 16)
        
        # 文件信息
        self.file_info_label = QLabel(get_text("Select a file to view details"))
        self.file_info_label.setWordWrap(True)
        details_layout.addWidget(self.file_info_label)
        
        # 预览区域
        self.preview_label = QLabel(get_text("Preview:"))
        details_layout.addWidget(self.preview_label)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(200)
        details_layout.addWidget(self.preview_text)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        
        self.open_file_btn = QPushButton(get_text("Open File"))
        self.open_folder_btn = QPushButton(get_text("Open Folder"))
        self.copy_path_btn = QPushButton(get_text("Copy Path"))
        
        # 设置按钮样式
        for btn in [self.open_file_btn, self.open_folder_btn, self.copy_path_btn]:
            btn.setMinimumHeight(32)
        
        button_layout.addWidget(self.open_file_btn)
        button_layout.addWidget(self.open_folder_btn)
        button_layout.addWidget(self.copy_path_btn)
        
        details_layout.addLayout(button_layout)
        details_layout.addStretch()
        
        splitter.addWidget(details_group)
        
    def create_status_bar(self):
        """创建状态栏"""
        self.status_bar = self.statusBar()
        self.status_label = QLabel(get_text("Ready"))
        self.status_bar.addWidget(self.status_label)
        
    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu(get_text("File"))
        
        self.open_action = QAction(get_text("Open Directory"), self)
        self.open_action.setShortcut(QKeySequence.StandardKey.Open)
        self.open_action.triggered.connect(self.browse_directory)
        file_menu.addAction(self.open_action)
        
        file_menu.addSeparator()
        
        self.exit_action = QAction(get_text("Exit"), self)
        self.exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        self.exit_action.triggered.connect(self.close)
        file_menu.addAction(self.exit_action)
        
        # 搜索菜单
        search_menu = menubar.addMenu(get_text("Search"))
        
        self.search_action = QAction(get_text("Start Search"), self)
        self.search_action.setShortcut(QKeySequence("Ctrl+F"))
        self.search_action.triggered.connect(self.start_search)
        search_menu.addAction(self.search_action)
        
        self.stop_action = QAction(get_text("Stop Search"), self)
        self.stop_action.setShortcut(QKeySequence("Ctrl+Shift+F"))
        self.stop_action.triggered.connect(self.stop_search)
        search_menu.addAction(self.stop_action)
        
        # 语言菜单
        language_menu = menubar.addMenu(get_text("Language"))
        
        self.chinese_action = QAction("中文", self)
        self.chinese_action.setCheckable(True)
        self.chinese_action.setChecked(True)
        self.chinese_action.triggered.connect(lambda: self.switch_language("chinese"))
        language_menu.addAction(self.chinese_action)
        
        self.english_action = QAction("English", self)
        self.english_action.setCheckable(True)
        self.english_action.setChecked(False)
        self.english_action.triggered.connect(lambda: self.switch_language("english"))
        language_menu.addAction(self.english_action)
        
        # 语言菜单不需要设置互斥，通过代码逻辑控制
        
        # 帮助菜单
        help_menu = menubar.addMenu(get_text("Help"))
        
        self.about_action = QAction(get_text("About"), self)
        self.about_action.triggered.connect(self.show_about)
        help_menu.addAction(self.about_action)
        
    def setup_connections(self):
        """设置信号连接"""
        self.browse_btn.clicked.connect(self.browse_directory)
        self.search_btn.clicked.connect(self.start_search)
        self.results_table.itemSelectionChanged.connect(self.on_file_selected)
        self.open_file_btn.clicked.connect(self.open_selected_file)
        self.open_folder_btn.clicked.connect(self.open_selected_folder)
        self.copy_path_btn.clicked.connect(self.copy_selected_path)
        
        # 关键字字段的回车键触发搜索
        self.keyword_edit.returnPressed.connect(self.start_search)
        
        # 目录选择框的回车键触发搜索
        self.dir_combo.lineEdit().returnPressed.connect(self.start_search)
        
    def browse_directory(self):
        """打开目录浏览器对话框"""
        current_dir = self.dir_combo.currentText() or os.path.expanduser("~")
        directory = QFileDialog.getExistingDirectory(
            self,
            get_text("Select Directory to Search"),
            current_dir
        )
        if directory:
            self.add_directory_to_history(directory)
            self.dir_combo.setCurrentText(directory)
            
    def start_search(self):
        """开始搜索过程"""
        directory = self.dir_combo.currentText().strip()
        keyword = self.keyword_edit.text().strip()
        
        if not directory:
            QMessageBox.warning(self, get_text("Warning"), get_text("Please select a directory to search."))
            return
            
        if not keyword:
            QMessageBox.warning(self, get_text("Warning"), get_text("Please enter a keyword to search."))
            return
            
        if not os.path.exists(directory):
            QMessageBox.warning(self, get_text("Warning"), get_text("Selected directory does not exist."))
            return
            
        # 将搜索目录添加到历史记录
        self.add_directory_to_history(directory)
            
        # 清理之前的搜索线程
        if self.search_thread and self.search_thread.isRunning():
            self.search_engine.stop_search()
            self.search_thread.quit()
            self.search_thread.wait()
            self.search_thread = None
        
        # 清除之前的结果
        self.results_table.clear()
        self.preview_text.clear()
        self.file_info_label.setText(get_text("Searching..."))
        
        # 更新UI状态
        self.search_btn.setEnabled(False)
        self.search_btn.setText(get_text("Searching..."))
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 不确定进度
        
        # 创建搜索引擎
        self.search_engine = SearchEngine()
        self.search_engine.set_search_params(
            directory=directory,
            keyword=keyword,
            case_sensitive=self.case_sensitive_cb.isChecked(),
            whole_word=self.whole_word_cb.isChecked(),
            include_subdirs=self.include_subdirs_cb.isChecked(),
            file_type=self.file_type_combo.currentText(),
            complete_search=self.complete_search_cb.isChecked()
        )
        
        # 连接信号
        self.search_engine.file_found.connect(self.on_file_found)
        self.search_engine.search_progress.connect(self.on_search_progress)
        self.search_engine.search_finished.connect(self.on_search_finished)
        self.search_engine.search_error.connect(self.on_search_error)
        
        # 在线程中启动搜索
        self.search_thread = QThread()
        self.search_engine.moveToThread(self.search_thread)
        self.search_thread.started.connect(self.search_engine.start_search)
        self.search_thread.start()
        
    def stop_search(self):
        """停止当前搜索"""
        if self.search_engine:
            self.search_engine.stop_search()
            
    def on_file_found(self, file_info):
        """处理找到包含关键字的文件"""
        logger.info(f"找到文件: {file_info['name']} - 匹配数: {file_info['matches']}")
        self.results_table.add_file(file_info)
        
    def on_search_progress(self, current, total):
        """处理搜索进度更新"""
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(current)
            
    def on_search_finished(self, total_files, found_files):
        """处理搜索完成"""
        self.search_btn.setEnabled(True)
        self.search_btn.setText(get_text("Search"))
        self.progress_bar.setVisible(False)
        
        # 清理线程
        if self.search_thread and self.search_thread.isRunning():
            self.search_thread.quit()
            self.search_thread.wait()
            self.search_thread = None
        
        self.status_label.setText(
            get_text("Search completed. Found {} files with keyword out of {} total files.").format(
                found_files, total_files
            )
        )
        
        if found_files == 0:
            self.file_info_label.setText(get_text("No files found containing the keyword."))
            
    def on_search_error(self, error_message):
        """处理搜索错误"""
        QMessageBox.critical(self, get_text("Search Error"), error_message)
        self.search_btn.setEnabled(True)
        self.search_btn.setText(get_text("Search"))
        self.progress_bar.setVisible(False)
        
        # 清理线程
        if self.search_thread and self.search_thread.isRunning():
            self.search_thread.quit()
            self.search_thread.wait()
            self.search_thread = None
        
    def on_file_selected(self):
        """处理结果表格中的文件选择"""
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            file_info = self.results_table.get_file_info(current_row)
            self.update_file_details(file_info)
            
    def update_file_details(self, file_info):
        """更新文件详情面板"""
        if not file_info:
            return
            
        # 更新文件信息
        info_text = f"""
<b>{get_text('File Name')}:</b> {file_info['name']}<br>
<b>{get_text('Path')}:</b> {file_info['path']}<br>
<b>{get_text('Size')}:</b> {file_info['size']}<br>
<b>{get_text('Modified')}:</b> {file_info['modified']}<br>
<b>{get_text('Keyword Matches')}:</b> {file_info['matches']}
"""
        self.file_info_label.setText(info_text)
        
        # 更新预览
        preview_text = file_info.get('preview', get_text('Preview not available'))
        self.preview_text.setPlainText(preview_text)
        
    def open_selected_file(self):
        """打开选中的文件"""
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            file_info = self.results_table.get_file_info(current_row)
            if file_info:
                self.results_table.open_file(file_info['path'])
                
    def open_selected_folder(self):
        """打开包含选中文件的文件夹"""
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            file_info = self.results_table.get_file_info(current_row)
            if file_info:
                self.results_table.open_folder(file_info['path'])
                
    def copy_selected_path(self):
        """复制选中文件的路径到剪贴板"""
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            file_info = self.results_table.get_file_info(current_row)
            if file_info:
                clipboard = QApplication.clipboard()
                clipboard.setText(file_info['path'])
                self.status_label.setText(get_text("Path copied to clipboard"))
                
    def show_about(self):
        """显示关于对话框"""
        about_text = f"""
{get_text("Excel Keyword Search Tool v1.0.0")}

{get_text("A professional tool for searching keywords in Excel files across directories.")}

{get_text("Features:")}
{get_text("• Fast multi-threaded search")}
{get_text("• Support for .xlsx and .xls files")}
{get_text("• Case-sensitive and whole word search")}
{get_text("• File preview and direct opening")}
{get_text("• Modern user interface")}

{get_text("© 2025 Professional Tools")}
        """
        QMessageBox.about(
            self,
            get_text("About Excel Keyword Search Tool"),
            about_text
        )
        
    def apply_styles(self):
        """应用明亮清晰的配色设计"""
        self.setStyleSheet("""
            /* 主窗口 - 浅灰背景 */
            QMainWindow {
                background-color: #f8f9fa;
            }
            
            /* 分组框 - 白色背景，深灰标题 */
            QGroupBox {
                font-weight: 600;
                font-size: 14px;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                margin-top: 16px;
                padding-top: 12px;
                background-color: #ffffff;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0px 6px 0px 6px;
                color: #495057;
                background-color: #ffffff;
                font-size: 15px;
            }
            
            /* 输入框 - 白色背景，深灰文字 */
            QLineEdit {
                padding: 12px 16px;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                font-size: 14px;
                color: #212529;
                selection-background-color: #0d6efd;
            }
            
            QLineEdit:focus {
                border: 2px solid #0d6efd;
                background-color: #ffffff;
            }
            
            QLineEdit::placeholder {
                color: #6c757d;
            }
            
            /* 按钮 - 蓝色主题 */
            QPushButton {
                padding: 8px 16px;
                border: none;
                border-radius: 6px;
                background-color: #0d6efd;
                color: #ffffff;
                font-weight: 500;
                font-size: 13px;
                min-height: 16px;
            }
            
            QPushButton:hover {
                background-color: #0b5ed7;
            }
            
            QPushButton:pressed {
                background-color: #0a58ca;
            }
            
            QPushButton:disabled {
                background-color: #e9ecef;
                color: #6c757d;
            }
            
            /* 表格 - 白色背景 */
            QTableWidget {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                gridline-color: #f8f9fa;
                selection-background-color: #0d6efd;
                selection-color: #ffffff;
                alternate-background-color: #f8f9fa;

            }
            
            QTableWidget::item {
                padding: 12px 16px;
                border: none;
                color: #212529;
                font-size: 13px;
                outline: none;  /* 禁用焦点框 */
            }
            
            QTableWidget::item:selected {
                background-color: #0d6efd;
                color: #ffffff;
            }
            
            QTableWidget::item:hover {
                background-color: #e9ecef;
            }
            
            QTableWidget::item:focus {
                outline: none;  /* 禁用焦点框 */
                border: none;   /* 禁用边框 */
            }
            
            /* 表头 - 透明背景 */
            QHeaderView::section {
                background-color: transparent;
                padding: 12px 16px;
                border: none;
                border-bottom: 2px solid #dee2e6;
                font-weight: 600;
                color: #495057;
                font-size: 12px;
            }
            
            /* 下拉框 - 白色背景 */
            QComboBox {
                padding: 12px 16px;
                padding-right: 40px;  /* 为下拉箭头留出空间 */
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                font-size: 14px;
                color: #212529;
                min-width: 120px;
            }
            
            QComboBox:focus {
                border: 2px solid #0d6efd;
            }
            
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #6c757d;
                margin-right: 8px;
                width: 12px;
                height: 12px;
            }
            
            /* 复选框 - 白色背景 */
            QCheckBox {
                spacing: 12px;
                font-size: 14px;
                color: #212529;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #dee2e6;
                border-radius: 4px;
                background-color: #ffffff;
            }
            
            QCheckBox::indicator:checked {
                background-color: #0d6efd;
                border-color: #0d6efd;
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOSIgdmlld0JveD0iMCAwIDEyIDkiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGQ9Ik0xIDQuNUw0LjUgOEwxMSAxIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPgo8L3N2Zz4K);
            }
            
            QCheckBox::indicator:hover {
                border-color: #0d6efd;
            }
            
            /* 文本编辑框 - 白色背景 */
            QTextEdit {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                padding: 12px 16px;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                font-size: 13px;
                color: #212529;
                line-height: 1.5;
            }
            
            /* 进度条 - 浅灰背景 */
            QProgressBar {
                border: 2px solid #dee2e6;
                border-radius: 4px;
                background-color: #f8f9fa;
                text-align: center;
                font-weight: 500;
                color: #495057;
                height: 8px;
            }
            
            QProgressBar::chunk {
                background-color: #0d6efd;
                border-radius: 2px;
            }
            
            /* 标签 - 深灰文字 */
            QLabel {
                color: #495057;
                font-size: 14px;
                font-weight: 400;
            }
            
            /* 菜单栏 - 白色背景 */
            QMenuBar {
                background-color: #ffffff;
                border-bottom: 2px solid #dee2e6;
                color: #495057;
                font-weight: 500;
                font-size: 13px;
            }
            
            QMenuBar::item {
                padding: 8px 16px;
                background-color: transparent;
            }
            
            QMenuBar::item:selected {
                background-color: #e9ecef;
                color: #495057;
            }
            
            /* 菜单 - 白色背景 */
            QMenu {
                background-color: #ffffff;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 8px 0px;
            }
            
            QMenu::item {
                padding: 8px 20px;
                color: #495057;
                font-size: 13px;
            }
            
            QMenu::item:selected {
                background-color: #e9ecef;
                color: #495057;
            }
            
            /* 状态栏 - 浅灰背景 */
            QStatusBar {
                background-color: #f8f9fa;
                border-top: 2px solid #dee2e6;
                color: #6c757d;
                font-size: 12px;
                padding: 8px 16px;
            }
            
            /* 分割器 - 灰色线条 */
            QSplitter::handle {
                background-color: #dee2e6;
                width: 2px;
            }
            
            QSplitter::handle:hover {
                background-color: #0d6efd;
            }
        """)
        
    def load_settings(self):
        """加载应用程序设置"""
        # 加载目录历史记录
        self.update_directory_history()
        
        # 设置上次使用的目录
        last_directory = self.settings.value('last_directory', '')
        if last_directory and os.path.exists(last_directory):
            self.dir_combo.setCurrentText(last_directory)
            
        self.keyword_edit.setText(self.settings.value('last_keyword', ''))
        self.case_sensitive_cb.setChecked(self.settings.value('case_sensitive', False, type=bool))
        self.whole_word_cb.setChecked(self.settings.value('whole_word', False, type=bool))
        self.include_subdirs_cb.setChecked(self.settings.value('include_subdirs', True, type=bool))
        self.complete_search_cb.setChecked(self.settings.value('complete_search', True, type=bool))
        
    def save_settings(self):
        """保存应用程序设置"""
        self.settings.setValue('last_directory', self.dir_combo.currentText())
        self.settings.setValue('last_keyword', self.keyword_edit.text())
        self.settings.setValue('case_sensitive', self.case_sensitive_cb.isChecked())
        self.settings.setValue('include_subdirs', self.include_subdirs_cb.isChecked())
        self.settings.setValue('complete_search', self.complete_search_cb.isChecked())
        
    def switch_language(self, language):
        """切换语言"""
        set_language(language)
        # 更新语言菜单项的选中状态
        self.chinese_action.setChecked(language == "chinese")
        self.english_action.setChecked(language == "english")
        # 立即更新UI语言
        self.update_ui_language()
    
    def update_ui_language(self):
        """更新UI语言"""
        # 更新窗口标题
        self.setWindowTitle(get_text("Excel Keyword Search Tool"))
        
        # GroupBox标题已删除，无需更新
        
        # 更新标签文本
        self.dir_label.setText(get_text("Search Directory:"))
        self.keyword_label.setText(get_text("Keyword:"))
        self.file_type_label.setText(get_text("File Types:"))
        
        # 更新按钮文本
        self.browse_btn.setText(get_text("Browse"))
        self.search_btn.setText(get_text("Search"))
        
        # 更新复选框文本
        self.case_sensitive_cb.setText(get_text("Case Sensitive"))
        self.whole_word_cb.setText(get_text("Whole Word"))
        self.include_subdirs_cb.setText(get_text("Include Subdirectories"))
        self.complete_search_cb.setText(get_text("Complete Search"))
        
        # 更新工具提示
        self.complete_search_cb.setToolTip(get_text("Complete Search Tooltip"))
        
        # 更新占位符文本
        self.dir_combo.setPlaceholderText(get_text("Select directory to search..."))
        self.keyword_edit.setPlaceholderText(get_text("Enter keyword to search..."))
        
        # 更新文件类型下拉框
        current_index = self.file_type_combo.currentIndex()
        self.file_type_combo.clear()
        self.file_type_combo.addItems([
            get_text("All Excel Files (.xlsx, .xls)"),
            get_text("Excel 2007+ (.xlsx)"),
            get_text("Excel 97-2003 (.xls)"),
            get_text("All Files")
        ])
        self.file_type_combo.setCurrentIndex(current_index)
        
        # 更新表格标题
        header_labels = [
            get_text("File Name"),
            get_text("Path"),
            get_text("Size"),
            get_text("Modified"),
            get_text("Keyword Matches")
        ]
        for i, label in enumerate(header_labels):
            self.results_table.horizontalHeaderItem(i).setText(label)
        
        # 更新菜单文本
        menubar = self.menuBar()
        for action in menubar.actions():
            # 根据当前语言更新菜单文本
            if action.text() in ["File", "文件"]:
                action.setText(get_text("File"))
            elif action.text() in ["Search", "搜索"]:
                action.setText(get_text("Search"))
            elif action.text() in ["Language", "语言"]:
                action.setText(get_text("Language"))
            elif action.text() in ["Help", "帮助"]:
                action.setText(get_text("Help"))
        
        # 更新菜单项文本
        self.open_action.setText(get_text("Open Directory"))
        self.exit_action.setText(get_text("Exit"))
        self.search_action.setText(get_text("Start Search"))
        self.stop_action.setText(get_text("Stop Search"))
        self.about_action.setText(get_text("About"))
        
        # 更新详情面板
        self.file_info_label.setText(get_text("Select a file to view details"))
        self.preview_label.setText(get_text("Preview:"))
        self.open_file_btn.setText(get_text("Open File"))
        self.open_folder_btn.setText(get_text("Open Folder"))
        self.copy_path_btn.setText(get_text("Copy Path"))
        
        # 更新状态栏
        self.status_label.setText(get_text("Ready"))
        
    def add_directory_to_history(self, directory):
        """添加目录到历史记录"""
        if not directory or not os.path.exists(directory):
            return
            
        # 获取当前历史记录
        history = self.settings.value('directory_history', [], type=list)
        
        # 如果目录已存在，先移除
        if directory in history:
            history.remove(directory)
            
        # 添加到开头
        history.insert(0, directory)
        
        # 限制历史记录数量（最多20个）
        history = history[:20]
        
        # 保存历史记录
        self.settings.setValue('directory_history', history)
        
        # 更新下拉框
        self.update_directory_history()
        
    def update_directory_history(self):
        """更新目录历史记录下拉框"""
        history = self.settings.value('directory_history', [], type=list)
        
        # 保存当前文本
        current_text = self.dir_combo.currentText()
        
        # 清空并重新添加历史记录
        self.dir_combo.clear()
        
        # 添加历史记录
        for directory in history:
            if os.path.exists(directory):
                self.dir_combo.addItem(directory)
                
        # 恢复当前文本
        if current_text:
            self.dir_combo.setCurrentText(current_text)
        
    def closeEvent(self, event):
        """处理应用程序关闭事件"""
        self.save_settings()
        # 注销语言切换回调
        unregister_language_change_callback(self.update_ui_language)
        if self.search_thread and self.search_thread.isRunning():
            self.stop_search()
            self.search_thread.quit()
            self.search_thread.wait()
        event.accept()
