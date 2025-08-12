@echo off
chcp 65001 >nul
echo 正在重置环境检查状态...
echo.

REM 删除环境检查标记文件
if exist ".env_checked" (
    del ".env_checked"
    echo ✓ 已删除环境检查标记文件
) else (
    echo - 未找到环境检查标记文件
)

echo.
echo 环境检查状态已重置！
echo 下次启动时会重新检查Python环境和依赖包。
echo.
pause
