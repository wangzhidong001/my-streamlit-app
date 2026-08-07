@echo off
chcp 65001 >nul
echo ========================================
echo   IDC 数据处理分析平台 启动中...
echo ========================================
echo.
echo 应用地址: http://localhost:8501
echo 按 Ctrl+C 停止应用
echo.
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501
pause
