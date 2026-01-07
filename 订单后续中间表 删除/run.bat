@echo off
chcp 65001 > nul
title NetSuite 订单后续中间表 批量删除

echo ==========================================
echo   NetSuite 订单后续中间表 批量删除工具
echo ==========================================
echo.

REM 切换到 bat 所在目录（支持中文 / 日文路径）
cd /d "%~dp0"

echo 当前目录：
echo %cd%
echo.

echo 正在启动 Python 程序……
echo.

python delete_middle_list.py

echo.
echo ==========================================
echo   程序已执行完毕
echo   如有问题请查看 logs 文件夹
echo ==========================================
echo.

pause
