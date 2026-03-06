@echo off
chcp 65001 >nul
echo ======================================================
echo   取证单号匹配工具 - 一键安装脚本
echo ======================================================
echo.

REM 获取脚本所在目录
set "SCRIPT_DIR=%~dp0"
set "TESSERACT_INSTALLER=%SCRIPT_DIR%tesseract-ocr-w64-setup-5.5.0.exe"
set "CHINESE_LANG_PACK=%SCRIPT_DIR%chi_sim.traineddata"

REM 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.7 或更高版本
    echo        下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/6] 检测 Python 版本...
python --version
echo.

REM 检查 pip 是否可用
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] pip 不可用，请重新安装 Python
    pause
    exit /b 1
)

echo [2/6] 安装 Python 依赖包...
echo 正在安装依赖，这可能需要几分钟...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查网络连接
    pause
    exit /b 1
)
echo 依赖安装完成！
echo.

REM 检查 Tesseract 是否安装
where tesseract >nul 2>&1
if %errorlevel% neq 0 (
    echo [3/6] Tesseract OCR 未安装，正在自动安装...
    echo.

    REM 检查本地是否有安装程序
    if not exist "%TESSERACT_INSTALLER%" (
        echo ======================================================
        echo   未找到 Tesseract 安装程序
        echo ======================================================
        echo.
        echo 请按以下步骤操作：
        echo.
        echo 1. 下载 Tesseract OCR 5.5.0：
        echo    https://github.com/UB-Mannheim/tesseract/releases
        echo    下载文件：tesseract-ocr-w64-setup-5.5.0.exe
        echo.
        echo 2. 将下载的安装程序复制到当前目录：
        echo    %SCRIPT_DIR%
        echo.
        echo 3. 重新运行此脚本
        echo.
        pause
        exit /b 1
    )

    echo 找到安装程序，正在安装 Tesseract OCR...
    echo 安装路径：D:\Tesseract-OCR
    echo.

    REM 静默安装 Tesseract
    "%TESSERACT_INSTALLER%" /S /D=D:\Tesseract-OCR
    if %errorlevel% neq 0 (
        echo [错误] Tesseract 安装失败
        pause
        exit /b 1
    )

    echo Tesseract OCR 安装完成！
    echo.

    REM 等待安装完成
    timeout /t 3 /nobreak >nul
) else (
    echo [3/6] Tesseract OCR 已安装...
    tesseract --version
    echo.
)

REM 检查并安装中文语言包
echo [4/6] 安装中文语言包...

REM 确定 Tesseract 安装路径
set "TESSERACT_PATH=D:\Tesseract-OCR"
if exist "C:\Program Files\Tesseract-OCR\tessdata" (
    set "TESSERACT_PATH=C:\Program Files\Tesseract-OCR"
)

REM 检查本地是否有中文语言包
if exist "%CHINESE_LANG_PACK%" (
    echo 复制中文语言包到 %TESSERACT_PATH%\tessdata...
    copy /Y "%CHINESE_LANG_PACK%" "%TESSERACT_PATH%\tessdata\" >nul
    if %errorlevel% equ 0 (
        echo 中文语言包安装成功！
    ) else (
        echo [警告] 中文语言包复制失败
    )
) else (
    echo 本地未找到中文语言包文件
    if exist "%TESSERACT_PATH%\tessdata\chi_sim.traineddata" (
        echo 已安装中文语言包
    ) else (
        echo [警告] 缺少中文语言包，OCR识别中文可能不可用
    )
)
echo.

REM 配置环境变量
echo [5/6] 配置环境变量...

REM 添加 Tesseract 到用户 PATH
setx PATH "%PATH%;%TESSERACT_PATH%" >nul 2>&1

REM 设置 TESSDATA_PREFIX
setx TESSDATA_PREFIX "%TESSERACT_PATH%\tessdata" >nul 2>&1

echo 环境变量已配置
echo 注意：环境变量将在新的命令行窗口中生效
echo.

echo [6/6] 安装完成！
echo.
echo ======================================================
echo   安装成功！现在可以运行工具了
echo ======================================================
echo.
echo 运行方法：
echo   1. 双击运行：启动工具.bat
echo   2. 或在命令行执行：streamlit run evidence_viewer.py
echo.
echo 工具会自动在浏览器中打开：http://localhost:8501
echo.
echo 重要提示：
echo   - 如果遇到命令找不到的错误，请重启命令行窗口或电脑
echo   - 首次运行 Streamlit 时可能需要输入邮箱（可直接跳过）
echo.
pause
