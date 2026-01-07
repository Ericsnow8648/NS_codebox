@echo off
chcp 65001 >nul

echo =========================================
echo RPA Akaden SETUP
echo =========================================
echo.

REM ---- Python 実行コマンドを探す（py 優先、なければ python） ----
set "PYCMD="

py --version >nul 2>nul
if %errorlevel%==0 (
    set "PYCMD=py"
)

if not defined PYCMD (
    python --version >nul 2>nul
    if %errorlevel%==0 (
        set "PYCMD=python"
    )
)

if not defined PYCMD (
    echo ERROR: Python not found.
    echo Please install Python from: https://www.python.org/downloads/
    pause
    exit /b
)

echo Found Python command: %PYCMD%
%PYCMD% --version
echo.

echo Updating pip...
%PYCMD% -m pip install --upgrade pip

echo Installing required packages...
%PYCMD% -m pip install selenium webdriver-manager pandas openpyxl

echo.
echo =========================================
echo SETUP COMPLETED
echo You can now run akaden.py
echo =========================================
pause
