@echo off
chcp 65001 >nul
echo ======================================================
echo   启动取证单号匹配工具
echo ======================================================
echo.

echo 正在启动...
echo 工具会在浏览器中自动打开：http://localhost:8501
echo.
echo 如需停止工具，请按 Ctrl+C
echo.

streamlit run evidence_viewer.py

pause
