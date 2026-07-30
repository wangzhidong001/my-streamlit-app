@echo off
chcp 65001 >nul
cd /d "C:\Users\ruijie\.trae-cn\work\6a4f8645e4b01f7722a32ca7"

echo ========================================
echo   Sync local code to Streamlit Cloud
echo ========================================
echo.

echo [1/4] Git status:
git status --short
echo.

echo [2/4] Enter commit message (empty = auto timestamp):
set /p MSG="> "

if "%MSG%"=="" (
    for /f "tokens=1-4 delims=/: " %%a in ("%date% %time%") do set TS=%%a%%b%%c%%d
    set MSG=auto-sync %TS%
)

echo.
echo [3/4] Committing...
git add .
git commit -m "%MSG%"

echo.
echo [4/4] Pushing to GitHub (Streamlit Cloud auto-redeploys)...
git push origin main

echo.
echo ========================================
echo   Done. Check Streamlit Cloud in 1-3 min
echo ========================================
pause
