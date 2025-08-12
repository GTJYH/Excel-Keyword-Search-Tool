#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel关键字搜索工具国际化支持模块
"""

# 中文翻译
CHINESE_TRANSLATIONS = {
    "Excel Keyword Search Tool": "Excel关键字搜索工具",
    "Search Configuration": "搜索配置",
    "Search Directory:": "搜索目录:",
    "Select directory to search...": "选择要搜索的目录...",
    "Select Directory to Search": "选择要搜索的目录",
    "Browse": "浏览",
    "Keyword:": "关键字:",
    "Enter keyword to search...": "输入要搜索的关键字...",
    "Case Sensitive": "区分大小写",
    "Whole Word": "完整单词",
    "Include Subdirectories": "包含子目录",
    "Complete Search": "完整搜索",
    "Complete Search Tooltip": "搜索文件的所有行（取消勾选只搜索前1000行以提高速度）",
    "File Types:": "文件类型:",
    "All Excel Files (.xlsx, .xls)": "所有Excel文件 (.xlsx, .xls)",
    "Excel 2007+ (.xlsx)": "Excel 2007+ (.xlsx)",
    "Excel 97-2003 (.xls)": "Excel 97-2003 (.xls)",
    "All Files": "所有文件",
    "Search": "搜索",
    "Searching...": "搜索中...",
    "Search Results": "搜索结果",
    "File Details": "文件详情",
    "Select a file to view details": "选择一个文件查看详情",
    "Preview:": "预览:",
    "Open File": "打开文件",
    "Open Folder": "打开文件夹",
    "Copy Path": "复制路径",
    "Ready": "就绪",
    "File": "文件",
    "Open Directory": "打开目录",
    "Exit": "退出",
    "Language": "语言",
    "Start Search": "开始搜索",
    "Stop Search": "停止搜索",
    "Help": "帮助",
    "About": "关于",
    "Warning": "警告",
    "Please select a directory to search.": "请选择要搜索的目录。",
    "Please enter a keyword to search.": "请输入要搜索的关键字。",
    "Selected directory does not exist.": "选择的目录不存在。",
    "Searching...": "搜索中...",
    "Search completed. Found {} files with keyword out of {} total files.": "搜索完成。在{}个文件中找到{}个包含关键字的文件。",
    "No files found containing the keyword.": "未找到包含关键字的文件。",
    "Search Error": "搜索错误",
    "File Name": "文件名",
    "Path": "路径",
    "Size": "大小",
    "Modified": "修改时间",
    "Keyword Matches": "关键字匹配",
    "Preview not available": "预览不可用",
    "Path copied to clipboard": "路径已复制到剪贴板",
    "About Excel Keyword Search Tool": "关于Excel关键字搜索工具",
    "Excel Keyword Search Tool v1.0.0": "Excel关键字搜索工具 v1.0.0",
    "A professional tool for searching keywords in Excel files across directories.": "一个用于在目录中搜索Excel文件关键字的专业工具。",
    "Features:": "功能特性:",
    "• Fast multi-threaded search": "• 快速多线程搜索",
    "• Support for .xlsx and .xls files": "• 支持 .xlsx 和 .xls 文件",
    "• Case-sensitive and whole word search": "• 区分大小写和完整单词搜索",
    "• File preview and direct opening": "• 文件预览和直接打开",
    "• Modern user interface": "• 现代化用户界面",
    "© 2025 Professional Tools": "© 2025 葛祥昇",
    "Failed to open file": "打开文件失败",
    "Failed to open folder": "打开文件夹失败",
    "Copy File Name": "复制文件名"
}

# 英文翻译（默认）
ENGLISH_TRANSLATIONS = {
    "Excel Keyword Search Tool": "Excel Keyword Search Tool",
    "Search Configuration": "Search Configuration",
    "Search Directory:": "Search Directory:",
    "Select directory to search...": "Select directory to search...",
    "Select Directory to Search": "Select Directory to Search",
    "Browse": "Browse",
    "Keyword:": "Keyword:",
    "Enter keyword to search...": "Enter keyword to search...",
    "Case Sensitive": "Case Sensitive",
    "Whole Word": "Whole Word",
    "Include Subdirectories": "Include Subdirectories",
    "Complete Search": "Complete Search",
    "Complete Search Tooltip": "Search all rows in files (uncheck to search only first 1000 rows for speed)",
    "File Types:": "File Types:",
    "All Excel Files (.xlsx, .xls)": "All Excel Files (.xlsx, .xls)",
    "Excel 2007+ (.xlsx)": "Excel 2007+ (.xlsx)",
    "Excel 97-2003 (.xls)": "Excel 97-2003 (.xls)",
    "All Files": "All Files",
    "Search": "Search",
    "Searching...": "Searching...",
    "Search Results": "Search Results",
    "File Details": "File Details",
    "Select a file to view details": "Select a file to view details",
    "Preview:": "Preview:",
    "Open File": "Open File",
    "Open Folder": "Open Folder",
    "Copy Path": "Copy Path",
    "Ready": "Ready",
    "File": "File",
    "Open Directory": "Open Directory",
    "Exit": "Exit",
    "Language": "Language",
    "Start Search": "Start Search",
    "Stop Search": "Stop Search",
    "Help": "Help",
    "About": "About",
    "Warning": "Warning",
    "Please select a directory to search.": "Please select a directory to search.",
    "Please enter a keyword to search.": "Please enter a keyword to search.",
    "Selected directory does not exist.": "Selected directory does not exist.",
    "Searching...": "Searching...",
    "Search completed. Found {} files with keyword out of {} total files.": "Search completed. Found {} files with keyword out of {} total files.",
    "No files found containing the keyword.": "No files found containing the keyword.",
    "Search Error": "Search Error",
    "File Name": "File Name",
    "Path": "Path",
    "Size": "Size",
    "Modified": "Modified",
    "Keyword Matches": "Keyword Matches",
    "Preview not available": "Preview not available",
    "Path copied to clipboard": "Path copied to clipboard",
    "About Excel Keyword Search Tool": "About Excel Keyword Search Tool",
    "Excel Keyword Search Tool v1.0.0": "Excel Keyword Search Tool v1.0.0",
    "A professional tool for searching keywords in Excel files across directories.": "A professional tool for searching keywords in Excel files across directories.",
    "Features:": "Features:",
    "• Fast multi-threaded search": "• Fast multi-threaded search",
    "• Support for .xlsx and .xls files": "• Support for .xlsx and .xls files",
    "• Case-sensitive and whole word search": "• Case-sensitive and whole word search",
    "• File preview and direct opening": "• File preview and direct opening",
    "• Modern user interface": "• Modern user interface",
    "© 2025 Professional Tools": "© 2025 Gexiangsheng",
    "Failed to open file": "Failed to open file",
    "Failed to open folder": "Failed to open folder",
    "Copy File Name": "Copy File Name"
}

# 当前语言（可以动态更改）
CURRENT_LANGUAGE = "chinese"

# 语言切换回调函数列表
_language_change_callbacks = []

def get_text(key):
    """获取指定键的翻译文本"""
    if CURRENT_LANGUAGE == "chinese":
        return CHINESE_TRANSLATIONS.get(key, key)
    else:
        return ENGLISH_TRANSLATIONS.get(key, key)

def set_language(language):
    """设置当前语言"""
    global CURRENT_LANGUAGE
    if CURRENT_LANGUAGE != language:
        CURRENT_LANGUAGE = language
        # 通知所有注册的回调函数
        for callback in _language_change_callbacks:
            try:
                callback()
            except Exception:
                pass

def register_language_change_callback(callback):
    """注册语言切换回调函数"""
    if callback not in _language_change_callbacks:
        _language_change_callbacks.append(callback)

def unregister_language_change_callback(callback):
    """注销语言切换回调函数"""
    if callback in _language_change_callbacks:
        _language_change_callbacks.remove(callback)
