@echo off
chcp 65001 >nul
echo 正在启动Excel关键字搜索工具...
echo.

REM 检查是否存在环境检查标记文件
if exist ".env_checked" (
    echo 环境已检查过，跳过检查...
    goto :start_app
)

REM 检查Python是否安装
echo 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.8或更高版本
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 检查依赖是否安装
echo 检查依赖包...
pip show PyQt6 >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖包...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo 错误：依赖包安装失败
        pause
        exit /b 1
    )
)

REM 创建环境检查标记文件
echo 环境检查完成 > .env_checked

:start_app
REM 启动程序
echo 启动程序...
python main.py

REM 如果程序异常退出，暂停显示错误信息
if errorlevel 1 (
    echo.
    echo 程序异常退出，请检查错误信息
    pause
)
