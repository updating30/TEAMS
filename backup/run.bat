@echo off
chcp 65001 > nul
echo ============================================
echo  Teams チャット → Excel 自動記入
echo ============================================
echo.
venv\Scripts\python main.py
echo.
pause
