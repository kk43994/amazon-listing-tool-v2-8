@echo off
title 亚马逊商品处理工具
echo ========================================
echo   亚马逊商品处理工具 v1.0
echo   Web界面启动中...
echo ========================================
echo.

cd /d "%~dp0"
python web\app.py

pause
