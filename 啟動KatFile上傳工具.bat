@echo off
chcp 65001 >nul
title KatFile 上傳工具啟動器

echo.
echo ========================================
echo    KatFile 上傳工具啟動器
echo ========================================
echo.

REM 檢查Python是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到Python
    echo.
    echo 請先安裝Python 3.7或更高版本：
    echo https://www.python.org/downloads/
    echo.
    echo 安裝時請確保勾選：
    echo • Add Python to PATH
    echo • Install tkinter
    echo.
    pause
    exit /b 1
)

echo ✅ 找到Python
python --version

echo.
echo 🔍 檢查依賴套件...

REM 嘗試啟動依賴檢查
python install_dependencies.py
if errorlevel 1 (
    echo.
    echo ❌ 依賴檢查失敗
    echo 請手動安裝必要套件：
    echo pip install requests python-docx py7zr
    echo.
    pause
    exit /b 1
)

echo.
echo 🚀 啟動主程式...
python start_katfile_uploader.py

if errorlevel 1 (
    echo.
    echo ❌ 程式啟動失敗
    echo 請檢查錯誤訊息並重試
    echo.
    pause
)

echo.
echo 👋 程式已結束
pause
