@echo off
chcp 65001 >nul 2>&1
title IDC 数据处理分析平台

echo ========================================
echo   IDC 数据处理分析平台 - 本地启动
echo ========================================
echo.

cd /d "%~dp0"

:: 检查 Python 是否可用
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 找不到 python，请先安装 Python 或确认已在 PATH 中。
    pause
    exit /b 1
)

:: 检查依赖是否已装
python -c "import streamlit, pandas, openpyxl, plotly, openai" >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装缺失的依赖，请稍候...
    pip install streamlit pandas openpyxl plotly openai -q
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请手动执行：
        echo   pip install streamlit pandas openpyxl plotly openai
        pause
        exit /b 1
    )
)

:: 检查 ZHIPU_API_KEY 环境变量（仅提醒）
if "%ZHIPU_API_KEY%"=="" (
    echo [提示] 未检测到 ZHIPU_API_KEY 环境变量
    echo        智能问答功能需要在页面中手动填写 API Key
    echo.
)

echo [就绪] 正在启动 Streamlit...
echo        浏览器将自动打开 http://localhost:8501
echo        按 Ctrl+C 可停止服务
echo.
python -m streamlit run app.py --server.port 8501

pause
