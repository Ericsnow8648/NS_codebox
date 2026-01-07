@echo off
REM ==========================
REM  强制 CMD 使用 UTF-8 显示中文
REM ==========================
chcp 65001 >nul

cd /d "%~dp0"

REM ---- 如果有虚拟环境，则激活 ----
if exist venv (
    call venv\Scripts\activate
)

echo [+] 正在启动 RPA 脚本...
python henpin_auto_akaden.py

echo.
echo [*] 脚本已结束，请按任意键关闭窗口...
pause >nul
