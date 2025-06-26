@echo off
chcp 65001 >nul
echo.
echo ========================================
echo   Excel MCP服务器配置生成器
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误: 未找到Python，请先安装Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ 检测到Python环境
echo.

REM 运行Python配置脚本
echo 🚀 启动配置生成器...
echo.
python setup_mcp_config.py

echo.
echo 按任意键退出...
pause >nul
