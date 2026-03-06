@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ======================================================
echo   取证单号匹配工具 - 便携版
echo ======================================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python
    echo.
    echo 本工具需要 Python 才能运行。
    echo 请访问 https://www.python.org/downloads/ 下载安装
    echo.
    pause
    exit /b 1
)

echo [1/3] 检查 Python...
python --version
echo.

echo [2/3] 检查依赖包...
pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装依赖包，请稍候...
    pip install -r requirements.txt -q
)
echo.

echo [3/3] 启动工具...
echo.
echo ======================================================
echo   工具即将在浏览器中打开
echo   请勿关闭此窗口
echo ======================================================
echo.

streamlit run evidence_viewer_portable.py

pause
